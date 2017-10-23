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
    Dim strYear, strEmp_no, strContr_date,strContr_rgst_no,strContr_Type,strAmt,strMainInsertFlag
    
    strYear = TRIM(Request("txtYear"))
    strEmp_no  = TRIM(Request("txtEmp_no"))
    strContr_date      = TRIM(Request("txtContr_date"))
    strContr_rgst_no   = TRIM(Request("txtContr_rgst_no"))
    strContr_Type      = TRIM(Request("txtContr_Type"))
    strAmt      = TRIM(Request("txtAmt"))   
    strMainInsertFlag  = TRIM(Request("txtMainInsertFlag"))     
%>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1307mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim isOpenPop
Dim strYear, strEmp_no, strContr_date,strContr_rgst_no,strContr_Type,strAmt, strMainInsertFlag

Dim gQuery_YN
Dim StartDate
dim fDiligAuth,fAuthCheck
<% StartDate	= GetSvrDate %>

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "I", "H") %>
End Sub
 
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
       
  	lgKeyStream = lgKeyStream & Trim(frm1.txtYear.value) & gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.txtContr_date.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtContr_rgst_no.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtContr_Type.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtContr_Cd.value) & gColSep 	
	lgKeyStream = lgKeyStream & Trim(frm1.txtAmt.value) & gColSep    
   
    lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
     lgKeyStream = lgKeyStream & Trim(frm1.txtSubmit_Cd.value) & gColSep    

	lgKeyStream = replace(lgKeyStream, "'", "''")
End Sub 
       
'========================================================================================================
' Name : InitComboBox()	
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCD
    
	iCodeArr = ""
	iNameArr = ""

    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		iDx = Replace(lgF0, Chr(11), "") +1
	end if
	iCodeArr = cdbl(idx) & Chr(11) & iCodeArr
	iNameArr = cdbl(idx) & Chr(11) & iNameArr
	   
    Call SetCombo2(frm1.txtYear, iCodeArr, iNameArr, Chr(11))
 '---------------------
    Call  CommonQueryRs(" REFERENCE , b.minor_nm "," b_configuration a join B_MINOR b on a.reference =b.minor_cd  "," a.MAJOR_CD ='H0126'  and b.major_cd = 'H0125' order by seq_no",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.txtContr_Type, iCodeArr, iNameArr, Chr(11))   
 '---------------------
    Call CommonQueryRs(" MINOR_CD,dbo.ufn_GetCodeName('H0126', MINOR_CD) "," b_configuration "," MAJOR_CD = 'H0126' order by SEQ_NO ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.txtContr_Cd, iCodeArr, iNameArr, Chr(11))   
    
    
     iCodeArr="Y" & Chr(11) &"N" & Chr(11)
     iNameArr="국세청자료" & Chr(11) &"그밖의자료" & Chr(11)
     
    Call SetCombo2(frm1.txtSubmit_Cd, iCodeArr, iNameArr, Chr(11))  
        
    
      
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = 4+1
    Grid1.SheetMaxrows = 3
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status

    strYear = "<%=strYear%>"
    If "<%=strEmp_no%>" <> "" Then
		strEmp_no = "<%=strEmp_no%>"
	Else
		strEmp_no = parent.txtEmp_no2.Value
	End IF
	strContr_date  = "<%=strContr_date%>"  
    strContr_rgst_no = "<%=strContr_rgst_no%>"
    strContr_Type = "<%=strContr_Type%>"
    strAmt = "<%=strAmt%>"
    strMainInsertFlag	= "<%=strMainInsertFlag%>"
        
    lgIntFlgMode = OPMD_CMODE   'insert mode
    gQuery_YN = ""              

    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

	call LoadInfTB19029()

    Call InitComboBox()
    Call LayerShowHide(0)

	Call parent.Click_OpenFrame(Replace(Ucase(BIZ_PGM_ID),"MB","MA"))

	if parent.txtName2.value ="" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if

	if  strContr_date <> "" then
	    Call SetToolBar("01110")

	    frm1.txtYear.value			= strYear
	    frm1.txtContr_date.value	= strContr_date
	    frm1.txtContr_rgst_no.value	= strContr_rgst_no
	    frm1.txtContr_Type.value	= strContr_Type	    
	    Call DbQuery(1)
	else    
	    Call SetToolBar("01010")

	    Call DbQuery(1)        
	end if
	frm1.txtSubmit_Cd.value ="N"
    Call LockField(Document)
End Sub

'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    Dim iret
    Err.Clear                                                                    '☜: Clear err status
	
    DbQuery = False                                                              '☜: Processing is NG

    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if    
    Call LayerShowHide(1)

    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&strMainInsertFlag="       & strMainInsertFlag

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    If gQuery_YN = "Y" Then                    '근태코드 아이템 체인지가 일어난경우.
		lgIntFlgMode = OPMD_CMODE              'create mode
    Else                                       '일일근태 현황에서 update를 하려고 들어 왔을때 
		lgIntFlgMode = OPMD_UMODE              'update mode
		ProtectTag(frm1.txtYear)
		ProtectTag(frm1.txtContr_date)
		ProtectTag(frm1.txtContr_rgst_no)	
		ProtectTag(frm1.txtContr_cd)	
		ProtectTag(frm1.txtContr_Type)		
    End if 
    
    Call SetToolBar("01110")	
	gQuery_YN = ""
	
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear

    lgIntFlgMode = OPMD_CMODE                'insert mode
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()

    Dim strDate
	Dim strVal

	On Error Resume Next
    Err.Clear                                                                    '☜: Clear err status

	DbSave = False														         '☜: Processing is NG

    if  frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtYear.focus()
        exit function
    end if

    if  Date_chk(frm1.txtContr_date.value, strDate) = True then
        frm1.txtContr_date.value = strDate
    else
  
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtContr_date.focus()
        exit function
    end if

    if  frm1.txtContr_rgst_no.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtContr_rgst_no.focus()
        exit function
    end if

    if  frm1.txtContr_Type.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtContr_Type.focus()
        exit function
    end if
    
    if  Trim(frm1.txtAmt.value) = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtAmt.focus()
        exit function
    Else
        if  Num_chk(frm1.txtAmt.value, strDate) = True then
           
            if mid(frm1.txtAmt.value,1,1) = "-" then
				Call DisplayMsgBox("800094","X","X","X")
				frm1.txtAmt.focus()
				exit function					
			end if
        else
            Call DisplayMsgBox("229924","X","X","X")
            frm1.txtAmt.focus()
            exit function
        end if
    end if
                
	If Not chkFieldLength(Document, "2") Then									         '☜: This function check required field
       Exit Function
    End If

    Call MakeKeyStream("C")

	Call LayerShowHide(1)

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
Function DbSaveOk()
	gQuery_YN = ""
	strMainInsertFlag = "N"	
    Call DbQuery(1)
End Function

'========================================================================================================
' Name : DbDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("C")
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
	frm1.txtYear.selectedIndex = 0    
    Call LockField(Document)	
    Call SetToolBar("01010")
	 frm1.txtSubmit_cd.value ="N"
    FncNew = True																 '☜: Processing is OK

End Function


'========================================================================================================
' Name : FncGetDiligAuth()
'========================================================================================================
Function FncGetDiligAuth(fDiligAuth,fAuthCheck)
    fDiligAuth = ""
    fAuthCheck = ""
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no = '" & Trim(parent.txtEmp_no.Value) & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub txtContr_Cd_OnChange()
    gQuery_YN = "Y"                                                    ' 아이템 체인지를 인식한다.전역변수 
    frm1.txtContr_Type.selectedIndex =frm1.txtContr_Cd.selectedIndex
End Sub

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
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
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
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
		                            <TD CLASS="ctrow01">정산연도</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtYear" ALT="정산연도" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01">기부일자</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" ID="form01" NAME="txtContr_date" TYPE="Text" MAXLENGTH=10 SiZE=10 alt="지급일자" tag="22D" ondblclick="VBScript:Call OpenCalendar('txtContr_date',3)">
		                            </TD>
                                </TR>                                
                                <TR>
		                            <TD CLASS="ctrow01">기부처사업자등록번호</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtContr_rgst_no" ALT="지급처사업자번호" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22"> '-' 제외
		                            </TD>
                                </TR>   
                                <TR>
		                            <TD CLASS="ctrow01">코드</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtContr_Cd" ALT="코드" STYLE="WIDTH: 220px" TAG="22"><OPTION VALUE=""></OPTION>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01">유형</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtContr_Type" ALT="유형" STYLE="WIDTH: 220px" TAG="24"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>제출구분</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtSubmit_cd" ALT="제출구분" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR> 
                                
                                
                                
		                        <TR>
		                            <TD CLASS="ctrow01">기부금액</TD>
		                            <TD CLASS="ctrow06"> <INPUT CLASS="form01" NAME="txtAmt" TYPE="Text" MAXLENGTH=13 SiZE=16 tag="22FU" style='TEXT-ALIGN: right'>
		                            </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>	

</BODY>
</HTML>