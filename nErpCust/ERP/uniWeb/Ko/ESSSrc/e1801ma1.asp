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

<%
    Dim emp_no,updateok
    emp_no = Trim(Request("emp_no"))
    updateok = Trim(Request("updateok"))
%>

<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1801mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim IsOpenPop
Dim emp_no
dim updateok

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep          
    else
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep 
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep         
    end if
End Sub  
      
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0120", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtpro_auth, iCodeArr, iNameArr, Chr(11))    

End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    frm1.txtemp_no1.focus() 

    emp_no = "<%=emp_no%>"
	updateok= "<%=updateok%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode

    Call InitComboBox()
    Call LockField(Document)	
    Call LayerShowHide(0)
    Call SetToolBar("01110")

    if  emp_no <> "" then
        frm1.txtemp_no1.value = emp_no
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
        Call DbQuery(1)
    end if
     'ProtectTag(frm1.txtuser_id)
     frm1.txtuser_id.classname = "form02"
     frm1.txtuser_id.readOnly = true
End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing

End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
	
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    lgIntFlgMode = OPMD_UMODE   'update mode

    if  frm1.txtdept_authv.value = "Y" then
        frm1.txtdept_auth1.checked = true
        frm1.txtdept_auth2.checked = false
    else
        frm1.txtdept_auth2.checked = true
        frm1.txtdept_auth1.checked = false
    end if

    if  frm1.txtuse_ynv.value = "Y" then
        frm1.txtuse_yn1.checked = true
        frm1.txtuse_yn2.checked = false
    else
        frm1.txtuse_yn2.checked = true
        frm1.txtuse_yn1.checked = false
    end if
    
	frm1.txtpassword.focus()
	'ProtectTag(frm1.txtemp_no1)
	'ProtectTag(frm1.txtuser_id)

    frm1.txtemp_no1.classname = "form02"
    frm1.txtemp_no1.readOnly = true

    frm1.txtuser_id.classname = "form02"
    frm1.txtuser_id.readOnly = true
	
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

    lgIntFlgMode = OPMD_CMODE   'insert mode
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	Dim lgStrSQL
	Dim lgObjConn
    Err.Clear                                                                    '☜: Clear err status

	With Frm1
	
		if .txtemp_no1.value = "" then
		   Call  DisplayMsgBox("800478","X","X","X")
		   .txtemp_no1.focus()
		   exit function
		end if
        if  frm1.txtpro_auth.value = "" then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtpro_auth.focus()
            exit function
        end if
        
        if  frm1.txtpassword.value = "" then
            Call  DisplayMsgBox("970021","X", frm1.txtpassword.alt,"X")
            frm1.txtpassword.focus()
            exit function
        end if

        if  frm1.txtdept_auth1.checked = true then
            frm1.txtdept_authv.value = "Y"
        else
            frm1.txtdept_authv.value = "N"
        end if

        if  frm1.txtuse_yn1.checked = true then
            frm1.txtuse_ynv.value = "Y"
        else
            frm1.txtuse_ynv.value = "N"
        end if
	End With

	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    if lgIntFlgMode <> OPMD_UMODE then
		if emp_no_check()=false then
			Call LayerShowHide(0)
			Exit Function
		end if
	end if

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key

        If lgIntFlgMode = OPMD_CMODE Then
           .txthpassword.value   = ConnectorControl.xCVTG(.txtpassword.value)
        Else
        
           If .txthpassword.value <> .txtpassword.value Then
              .txthpassword.value   = ConnectorControl.xCVTG(.txtpassword.value)
           End If
        
        End If   

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk()
    Call DbQuery(1)
End Function

'========================================================================================================
' Name : DbDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
	Dim IntRetCD
	
    Err.Clear                                                                    '☜: Clear err status
	
	if  frm1.txtemp_no1.value = "" then
       Call  DisplayMsgBox("800478","X","X","X")
       frm1.txtemp_no1.focus()
       exit function
    end if
	
	IntRetCD =  DisplayMsgBox("900003", VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	DbDelete = False			                                                 '☜: Processing is NG

	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

	With Frm1
		.txtMode.value        = "UID_M0003"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                              '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
'========================================================================================================
Function DbDeleteOk()
	Call FncNew()	
End Function

'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	lgIntFlgMode = OPMD_CMODE              'create mode

    Call ClearField(document,2)
    Call LockField(Document)
 
    Call SetToolbar("0111")
    frm1.txtdept_auth1.checked = true
    frm1.txtuse_yn1.checked = true
    frm1.txtemp_no1.focus()    
    'ProtectTag(frm1.txtuser_id)    
    frm1.txtuser_id.classname = "form02"
    frm1.txtuser_id.readOnly = true

    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : OpenEmp()
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)
	
	if lgIntFlgMode = OPMD_UMODE then
		Exit Function
	end if

	If IsOpenPop = True Then Exit Function
	If frm1.txtemp_no1.className = "ted" Then Exit Function
	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no1.value			' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
    arrParam(2) = Trim(parent.txtinternal_cd.Value)'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa2.asp", Array(arrParam), _
		"dialogWidth=615px; dialogHeight=413px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtemp_no1.value = arrRet(0)
	    frm1.txtname1.value = arrRet(1)
	    frm1.txtuser_id.value = arrRet(2)
		frm1.txtpassword.focus()
	End If	
			
End Function

'========================================================================================================
'   Event Name : emp_no_check()  
'========================================================================================================

Function emp_no_check()
	dim strEmp_no,IntRetCd,strUID,IntRetCd2
	dim tem_name1,tem_user_id
	
 	strEmp_no = Trim(frm1.txtemp_no1.value)
 	strUID = Trim(frm1.txtuser_id.value)
    IntRetCd = CommonQueryRs(" NAME,res_no "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	tem_name1 = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))
	tem_user_id=Trim(Replace(ConvSPChars(lgF1),Chr(11),""))    
	
    If IntRetCd = false then
		Call DisplayMsgbox("800048","X","X","X")
		frm1.txtuser_id.value=""
        frm1.txtname1.value = ""
        frm1.txtemp_no1.focus
        emp_no_check = false

    ELSE    
	    IntRetCd2 = CommonQueryRs(" emp_no "," E11002T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		if IntRetCd2=true then
			Call DisplayMsgbox("800473","X","X","X")
			frm1.txtuser_id.value=""
			frm1.txtname1.value = ""
			frm1.txtemp_no1.focus			
			emp_no_check = false
		else
			frm1.txtname1.value = tem_name1
			frm1.txtuser_id.value= tem_user_id
			emp_no_check = true
        end if
    END IF
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Function txtemp_no1_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no

	frm1.txtemp_no1.value=Trim(frm1.txtemp_no1.value)
    IF Trim(frm1.txtemp_no1.value) = "" THEN
        frm1.txtname1.value = ""
        txtemp_no1_Onchange = true
    ELSE
		strEmp_no = Trim(frm1.txtemp_no1.value)
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
            frm1.txtname1.value = ""
            frm1.txtemp_no1.focus
        ELSE    
            frm1.txtname1.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '사번에 해당하는 이름 
            txtemp_no1_Onchange = true
        END IF
    END IF 
End Function

Function Document_onKeyDown()
	Dim CuEvObj,KeyCode
	
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	Document_onKeyDown	= True	
End Function

Function txtemp_no1_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
			if emp_no_check() then
				txtemp_no1_onKeyDown	= True
			else 
				txtemp_no1_onKeyDown	= False
			end if
	End Select		
	txtemp_no1_onKeyDown	= True	
End Function

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=732>
        <TR>
           <td height="10"></td>
        </TR>
        <TR>
            <TD>
                <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
		            <TR>
		                <TD CLASS="ctrow01">사번</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <INPUT CLASS="form01" NAME="txtemp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22">&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtemp_no1.value)">
		                    <INPUT CLASS="form02" NAME="txtname1" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01">사용자ID</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <INPUT CLASS="form02" NAME="txtuser_id" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22">
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01">패스워드</TD>
		                <TD CLASS="ctrow06" colspan=3><INPUT CLASS="form01" NAME="txtpassword" TYPE="password" MAXLENGTH=10 alt="패스워드" SiZE=10 tag="22">
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01">레벨</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtpro_auth" STYLE="WIDTH: 100px" tag="22"><OPTION VALUE=""></OPTION></SELECT>
		                    사용가능한 프로그램의 그룹을 지정합니다.
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01">자료권한</TD>
		                <TD CLASS="ctrow06" colspan=3>
		               	    <INPUT CLASS="ftgray" TYPE="RADIO" NAME="txtdept_auth" tag="2" CHECKED ID="txtdept_auth1" VALUE='Y'><LABEL FOR="txtdept_auth1">사용</LABEL>
    			            <INPUT CLASS="ftgray" TYPE="RADIO" NAME="txtdept_auth" tag="2" ID="txtdept_auth2" VALUE='N'><LABEL FOR="txtdept_auth2">미사용</LABEL>
                            <INPUT CLASS="ftgray" TYPE=HIDDEN NAME="txtdept_authv">
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01">사용여부</TD>
		                <TD CLASS="ctrow06" colspan=3>
		               	    <INPUT CLASS="ftgray" TYPE="RADIO" NAME="txtuse_yn" tag="2" CHECKED ID="txtuse_yn1" VALUE=1><LABEL FOR="txtuse_yn1">사용</LABEL>
    			            <INPUT CLASS="ftgray" TYPE="RADIO" NAME="txtuse_yn" tag="2" ID="txtuse_yn2" VALUE=2><LABEL FOR="txtuse_yn2">미사용</LABEL>
                            <INPUT CLASS="ftgray" TYPE=HIDDEN NAME="txtuse_ynv">
                        </TD>      
		            </TR>
		            <TR height=110>
		                <TD CLASS="ctrow01"></TD>
		                <TD CLASS="ctrow06" colspan=3></TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
           <td height="10"></td>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 width=700 HEIGHT=0 border=0 bgcolor=#ffffff>
        <TR><TD HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txthpassword">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
</FORM>	

</BODY>
</HTML>
