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
    Dim trip_strt_dt
    Dim trip_cd
    Dim app_yn

    trip_strt_dt = Trim(Request("trip_strt_dt"))
    trip_cd = Trim(Request("trip_cd"))
    app_yn = Trim(Request("app_yn"))
    
%>


<Script Language="VBScript">
Option Explicit  

Const BIZ_PGM_ID      = "e1701mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim trip_strt_dt
Dim trip_cd
Dim app_yn
dim fDiligAuth,fAuthCheck
Dim StartDate

<%StartDate	= GetSvrDate%>

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
        lgKeyStream = lgKeyStream & Trim(frm1.txtTrip_strt_dt.value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtTrip_end_dt.value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.trip_cd.value) & gColSep
        lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
        lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
        lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep             

End Sub  
      
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
	Call CommonQueryRs(" dilig_cd, dilig_nm "," hca010t ", " dilig_cd in (" & FilterVar("98", "''", "S") & "," & FilterVar("99", "''", "S") & ") " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txttrip_cd, iCodeArr, iNameArr,Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
  
    trip_strt_dt = "<%=trip_strt_dt%>"
    trip_cd = "<%=trip_cd%>"
    app_yn="<%=app_yn%>"
	frm1.trip_cd.value = trip_cd
    lgIntFlgMode = OPMD_CMODE   'insert mode

    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

    Call InitComboBox()
    Call LayerShowHide(0)

    if  trip_strt_dt <> "" then
        frm1.txttrip_strt_dt.value = trip_strt_dt
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
    else
        frm1.txttrip_strt_dt.value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
        frm1.txttrip_end_dt.value =  frm1.txttrip_strt_dt.value
        frm1.txtTrip_amt.value = 0
    end if
   
    Call LockField(Document)

	If trip_cd = "" and trip_strt_dt = "" Then

	       Call SetToolBar("01010")
	End if

    if  trip_cd <> "" then
        frm1.txttrip_cd.value = UNIDateClientFormat(trip_cd)
       
        Call DbQuery(1)
    else
        Call DbQueryEmp(1)
    end if

	frm1.txtAppyn.value = "<%=Trim(Request("txtapp_yn"))%>"

End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	With Frm1

        if  Date_chk(.txttrip_strt_dt.value, strDate) = True then
            .txttrip_strt_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txttrip_strt_dt.focus()
            exit function
        end if

        if  Date_chk(.txttrip_strt_dt.value, strDate) = True then
            .txttrip_end_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txttrip_end_dt.focus()
            exit function
        end if
   
    End With
	If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    DbQuery = False                                                              '☜: Processing is NG

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryEmp
'========================================================================================
Function DbQueryEmp(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQueryEmp = False                                                              '☜: Processing is NG
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")


    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQueryEmp = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    lgIntFlgMode = OPMD_UMODE   'update mode

    ProtectTag(frm1.txttrip_strt_dt)
    ProtectTag(frm1.txttrip_end_dt)
    ProtectTag(frm1.txtTrip_cd)

    frm1.txtTrip_cd.disabled = true
    
    if app_yn  = "Y" or  app_yn  = "R" then
	 	ProtectTag(frm1.txtTrip_loc)
	 	ProtectTag(frm1.txtTrip_amt)
	 	ProtectTag(frm1.txtremark)
	 	ProtectTag(frm1.txtApp_name)
	 	ProtectTag(frm1.txtApp_emp_no)       
	     Call SetToolBar("01000")
	 else 	
	 	Call SetToolBar("01110")
	end if		    
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    lgIntFlgMode = OPMD_CMODE   'insert mode
    Call ClearField(document,2)
        frm1.txttrip_strt_dt.value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
        frm1.txttrip_end_dt.value =  frm1.txttrip_strt_dt.value
    frm1.txtTrip_amt.value = 0
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	Dim strAppyn
	
    Err.Clear                                                                    '☜: Clear err status
	'----반려처리된 사항에 대해서는 수정이 이루어질 수 없다.
	strAppyn = frm1.txtAppyn.value 
	If strAppyn = "R" Then
		Call DisplayMsgBox("800477","X","X","X")
        exit function
	End IF

	With Frm1

        if .txtEmp_no.value = "" then
            Call DisplayMsgBox("800006","X","X","X")
            .txtTrip_strt_dt.focus()
            exit function
        end if

        if .txtEmp_no.value = .txtApp_emp_no.value  then
            Call DisplayMsgBox("800476","X","X","X")
		        .txtapp_emp_no.value    = ""
		        .txtapp_name.value      = ""
		        .txtApp_emp_no.focus()
            exit function
        end if
        
        if  Date_chk(.txtTrip_strt_dt.value, strDate) = True then
            .txtTrip_strt_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtTrip_strt_dt.focus()
            exit function
        end if

        if  .txtTrip_cd.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtTrip_cd.focus()
            exit function
        Else
			.trip_cd.value = .txtTrip_cd.value
        end if

        if  Date_chk(.txtTrip_end_dt.value, strDate) = True then
            .txtTrip_end_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtTrip_end_dt.focus()
            exit function
        end if

		If CompareDateByFormat(.txttrip_strt_dt.value,.txttrip_end_dt.value,.txttrip_strt_dt.Alt,.txttrip_end_dt.Alt,"970025", gDateFormat, gComDateType,TRUE) = False Then
            .txtTrip_strt_dt.focus()
            exit function
		END IF

        if  Trim(.txtTrip_loc.value) = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtTrip_loc.focus()
            exit function
        end if

        if  Trim(.txtTrip_amt.value) = "" then
			.txtTrip_amt.value = 0
        Else
            if  Num_chk(.txtTrip_amt.value, strDate) = True then
           
                if mid(.txtTrip_amt.value,1,1) = "-" then
					Call DisplayMsgBox("800094","X","X","X")
					.txtTrip_amt.focus()
					exit function					
				end if
            else
                Call DisplayMsgBox("229924","X","X","X")
                .txtTrip_amt.focus()
                exit function
            end if
        end if

        if  Trim(.txtApp_emp_no.value) = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtApp_emp_no.focus()
            exit function
        end if

	End With
	
	If Not chkFieldLength(Document, "2") Then									         '☜: This function check required field
       Exit Function
    End If

	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("S")
    
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk(pCd)
	frm1.trip_cd.value = pCd
    Call DbQuery(1)
End Function

'========================================================================================================
' Function Name : DbSaveFail
'========================================================================================================
Function DbSaveFail()
End Function

'========================================================================================================
' Name : DbDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)

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
	lgIntFlgMode = OPMD_CMODE
	Call FncNew()	
End Function

'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    lgIntFlgMode = OPMD_CMODE   'insert mode

    Call ClearField(document,2)

    Call SetToolbar("01010")

    Call LockField(Document)	

    frm1.txttrip_strt_dt.value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
    frm1.txttrip_end_dt.value =  frm1.txttrip_strt_dt.value
    frm1.txtTrip_amt.value = 0
    frm1.txtAppyn.value = ""
    frm1.txttrip_strt_dt.focus()    
    app_yn = ""

    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : OpenEmp()
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)

	If OpenEmp = True or  app_yn = "Y" or  app_yn = "R" Then Exit Function

	OpenEmp = True

	arrParam(0) = frm1.txtApp_Emp_no.value			' Code Condition
	arrParam(1) = ""								' Name Cindition
    arrParam(2) = "APPROVAL_CODE"					'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa3.asp", Array(arrParam), _
		"dialogWidth=546px; dialogHeight=387px; center: Yes; help: No; resizable: No; status: No;")
		
	OpenEmp = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtApp_emp_no.value = arrRet(0)
	    frm1.txtApp_name.value = arrRet(1)
	End If	

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

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Function Document_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	Document_onKeyDown	= True	
End Function

Function txtApp_emp_no_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no

	frm1.txtApp_emp_no.value=Trim(frm1.txtApp_emp_no.value)
    IF Trim(frm1.txtApp_emp_no.value) = "" THEN
        frm1.txtApp_name.value = ""
        txtApp_emp_no_Onchange = true
    ELSE
		strEmp_no = Trim(frm1.txtApp_emp_no.value)
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
            frm1.txtApp_name.value = ""
            frm1.txtApp_emp_no.focus
        ELSE    
            frm1.txtApp_name.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '사번에 해당하는 이름 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
End Function
Function txtEmp_no2_Onchange()
        Call DbQueryEmp(1)	
End Function

Function txtApp_emp_no_onKeyDown()
Dim CuEvObj,KeyCode,IntRetCd
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	txtApp_emp_no_onKeyDown	= True	
End Function

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>

                    <TR>
                        <TD>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>출장기간</TD>
		                            <TD CLASS="ctrow06" colspan=3>
		                                <INPUT CLASS="form01" ID="txtTrip_strt_dt" NAME="txtTrip_strt_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 Alt="출장시작일" TAG="22" ondblclick="VBScript:Call OpenCalendar('txtTrip_strt_dt',3)" >&nbsp;~&nbsp;
		                                <INPUT CLASS="form01" ID="txtTrip_end_dt" NAME="txtTrip_end_dt" TYPE="Text" MAXLENGTH=10 SiZE=10   Alt="출장종료일" TAG="22" ondblclick="VBScript:Call OpenCalendar('txtTrip_end_dt',3)" >&nbsp;
	                                
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>출장</TD>
		                            <TD CLASS="ctrow06" colspan=3><SELECT CLASS="form01" NAME="txtTrip_cd" STYLE="WIDTH: 150px" tag="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>출장내용</TD>
		                            <TD CLASS="ctrow06" colspan=3>
		                                <INPUT CLASS="form01" NAME="txtTrip_loc" ALT ="출장내용" TYPE="Text" MAXLENGTH=50 SiZE=40 tag="22">
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>출장비</TD>
		                            <TD CLASS="ctrow06" colspan=3>
		                                <INPUT CLASS="form01" NAME="txtTrip_amt" TYPE="Text" MAXLENGTH=13 SiZE=16 style='TEXT-ALIGN: right'  tag="22">
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>비고</TD>
		                            <TD CLASS="ctrow06" colspan=3>
		                                <INPUT CLASS="form01" NAME="txtremark" alt="비고" TYPE="Text" MAXLENGTH=50 SiZE=50  tag="22">
                                    </TD>      
		                        </TR>
		                        <TR>
		                            <TD CLASS="ctrow01" NOWRAP>승인자</TD>
		                            <TD CLASS="ctrow06" colspan=3><INPUT CLASS="form01" NAME="txtApp_emp_no" ALT="승인사번" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22">&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtApp_emp_no.value)">
		                                    <INPUT CLASS="form01" NAME="txtApp_name" TYPE="Text" MAXLENGTH=20 SiZE=20  tag="24">
		                            </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
           <td height="10"></td>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
    <INPUT TYPE=HIDDEN NAME="trip_cd">
    <INPUT TYPE=HIDDEN NAME="txtAppyn">
</FORM>

</BODY>
</HTML>
