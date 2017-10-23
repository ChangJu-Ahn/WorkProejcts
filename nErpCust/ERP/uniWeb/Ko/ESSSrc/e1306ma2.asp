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
    Dim strYear, strEmp_no, strMed_date, strMed_name,strFamily_res_no, strFamily_name,strMainInsertFlag
    dim strMed_Resno
    
    strYear				= TRIM(Request("txtYear"))
    strEmp_no			= TRIM(Request("txtEmp_no"))
    strMed_date			= TRIM(Request("txtMed_date"))
    strMed_name			= TRIM(Request("txtMed_name"))
    strMed_Resno		= TRIM(Request("txtMed_Resno"))
    strFamily_res_no	= TRIM(Request("txtFamily_res_no"))  
    strFamily_name		= TRIM(Request("txtFamily_name")) 
    strMainInsertFlag   = TRIM(Request("txtMainInsertFlag")) 

%>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1306mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim isOpenPop
Dim strYear, strEmp_no, strMed_date, strMed_name,strFamily_res_no,strFamily_name, strMainInsertFlag
dim strMed_Resno
Dim gQuery_YN
dim fDiligAuth,fAuthCheck

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
  	
	'lgKeyStream = lgKeyStream & UniConvDateAToB(Trim(frm1.txtMed_date.value) ,gDateFormat, gServerDateFormat) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtMed_date.value)& gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtMed_name.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtMed_rgst_no.value) & gColSep 	
	lgKeyStream = lgKeyStream & Trim(frm1.txtFamily_Name.value) & gColSep 	
	lgKeyStream = lgKeyStream & Trim(frm1.txtRel_cd.value) & gColSep 	
	lgKeyStream = lgKeyStream & Trim(frm1.txtFamily_res_no.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtType_cd.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtAmt.value) & gColSep    
	lgKeyStream = lgKeyStream & Trim(frm1.txtMed_text.value) & gColSep 
    lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    lgKeyStream = lgKeyStream & Trim(frm1.txtSubmit_Cd.value) & gColSep 
    lgKeyStream = lgKeyStream & Trim(frm1.txtProvcnt.value) & gColSep     
 	
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
    
     iCodeArr="Y" & Chr(11) &"N" & Chr(11)
     iNameArr="국세청자료" & Chr(11) &"그밖의자료" & Chr(11)
     
    Call SetCombo2(frm1.txtSubmit_Cd, iCodeArr, iNameArr, Chr(11))  
        
    
    
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

	strMed_date			= "<%=strMed_date%>"  
    strMed_name			= "<%=strMed_name%>"
    strFamily_res_no	= "<%=strFamily_res_no%>"
    strFamily_name		= "<%=strFamily_name%>"
    strMed_Resno		= "<%=strMed_Resno%>"
    strMainInsertFlag	= "<%=strMainInsertFlag%>"
    

 
	ProtectTag(frm1.txtRel_nm)	 
	ProtectTag(frm1.txtFamily_res_no)
	ProtectTag(frm1.txtType)   
	

    lgIntFlgMode = OPMD_CMODE   'insert mode
    gQuery_YN = ""              

    call FncGetDiligAuth(fDiligAuth,fAuthCheck)

    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

    Call InitComboBox()
    Call LayerShowHide(0)
 
	Call parent.Click_OpenFrame(Replace(Ucase(BIZ_PGM_ID),"MB","MA"))

	if parent.txtName2.value ="" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if
	frm1.txtSubmit_Cd.value ="N"
	if  strFamily_res_no <> "" then
	    Call SetToolBar("01110")

	    frm1.txtYear.value		= strYear
	    frm1.txtMed_date.value	= strMed_date
	    frm1.txtMed_name.value	= strMed_name
	    frm1.txtMed_rgst_no.value	= strMed_ResNo
	    frm1.txtFamily_Name.value	=  strFamily_name
	    frm1.txtSubmit_cd.value		= "<%=Request("txtflag")%>" 
	    frm1.txtFamily_res_no.value	= strFamily_res_no

	    Call DbQuery(1)
	else    
	    Call SetToolBar("01010")

	    Call DbQuery(1)        
	end if

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
		ProtectTag(frm1.txtMed_date)
		'ProtectTag(frm1.txtMed_name)	
		ProtectTag(frm1.txtFamily_Name)	
		ProtectTag(frm1.txtRel_nm)	
		ProtectTag(frm1.txtFamily_res_no)	
		ProtectTag(frm1.txtType)
		ProtectTag(frm1.txtMed_rgst_no)	
		ProtectTag(frm1.txtSubmit_cd)	
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
	Dim strAppyn

	On Error Resume Next
    Err.Clear                                                                    '☜: Clear err status

	DbSave = False														         '☜: Processing is NG
    
    if  frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtYear.focus()
        exit function
    end if

    if  Date_chk(frm1.txtMed_date.value, strDate) = True then
        frm1.txtMed_date.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMed_date.focus()
        exit function
    end if

    if  frm1.txtMed_name.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMed_name.focus()
        exit function
    end if

    if  frm1.txtFamily_Name.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtFamily_Name.focus()
        exit function
    end if

    If txtFamily_Name_Onchange() Then      
        Exit Function
    End if
  
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
	ProtectTag(frm1.txtRel_nm)	 
	ProtectTag(frm1.txtFamily_res_no)
	ProtectTag(frm1.txtType)   
	frm1.txtSubmit_cd.value ="N"
	frm1.txtprovcnt.value ="1"
    Call SetToolBar("01010")

    FncNew = True																 '☜: Processing is OK

End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub txtYear_OnChange()
    gQuery_YN = "Y"                                                    ' 아이템 체인지를 인식한다.전역변수 
End Sub

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
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no = '" & Trim(parent.txtEmp_no.Value) & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function

'========================================================================================================
' Name : OpenCode()
'========================================================================================================
Function OpenCode(pCode)
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HFA150T"  								' TABLE
	arrParam(1) = "family_name"								' Code Condition
	arrParam(2) = "dbo.ufn_GetCodeName('H0140',FAMILY_REL) "								' Name Cindition
	arrParam(3) =  frm1.txtFamily_Name.value					' Code값 
	arrParam(4) = " MEDI_YN='Y' and year_yy = " & FilterVar(frm1.txtYear.value, "''", "S") & " and emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " and family_name like " 	'WHERE조건 

	arrRet = window.showModalDialog("e1codepopa1.asp", Array(arrParam), _
		"dialogWidth=546px; dialogHeight=340px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	
 
		
	If arrRet(0) = "" Then
		Exit Function
	Else
 		frm1.txtFamily_Name.value = arrRet(0)

 		Call CommonQueryRs("  FAMILY_RES_NO, FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " &_
		"  ,CASE 	WHEN PARIA_YN ='Y' THEN 'A'  "&_
		" 		WHEN   "&frm1.txtYear.value &" -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN 'B' "&_
		" 		ELSE '' END "&_
		"  ,CASE 	WHEN PARIA_YN ='Y' THEN '장애자'  "&_
		" 		WHEN    "&frm1.txtYear.value &" -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN '경로자' "&_
		" 		ELSE '' END ",_
		"  HFA150T ",_
		" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND YEAR_YY = " & FilterVar(frm1.txtYear.value, "''", "S") & " AND FAMILY_NAME= " & FilterVar(frm1.txtFamily_Name.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		frm1.txtFamily_res_no.value = Trim(Replace(lgF0,Chr(11),""))
		frm1.txtRel_cd.value		= Trim(Replace(lgF1,Chr(11),""))
		frm1.txtRel_nm.value		= Trim(Replace(lgF2,Chr(11),""))
		frm1.txtType_cd.value		= Trim(Replace(lgF3,Chr(11),""))
		frm1.txtType.value			= Trim(Replace(lgF4,Chr(11),"")) 

 	End If	
 
End Function

'========================================================================================================
'   Event Name :txtFamily_Name_Onchange    
'========================================================================================================
Function txtFamily_Name_Onchange()
    Dim IntRetCd
 
    If frm1.txtEmp_no.value = "" Then
		frm1.txtRel_cd.value = ""
		frm1.txtRel_nm.value = ""
		frm1.txtFamily_res_no.value = ""
		frm1.txtType.value = ""
		frm1.txtType_cd.value = ""
    Else

 		IntRetCd = CommonQueryRs("  FAMILY_RES_NO, FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " &_
		"  ,CASE 	WHEN PARIA_YN ='Y' THEN 'A'  "&_
		" 		WHEN   "&frm1.txtYear.value &" -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN 'B' "&_
		" 		ELSE '' END "&_
		"  ,CASE 	WHEN PARIA_YN ='Y' THEN '장애자'  "&_
		" 		WHEN   "&frm1.txtYear.value &" -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN '경로자' "&_
		" 		ELSE '' END ",_
		"  HFA150T ",_
		" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND MEDI_YN = 'Y' AND YEAR_YY = " & FilterVar(frm1.txtYear.value, "''", "S") & " AND FAMILY_NAME= " & FilterVar(frm1.txtFamily_Name.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
	    if  IntRetCd = False then
   			Call  DisplayMsgBox("971012","X","부양가족공제자 의료비 항목에 체크가 되어 있는지","X")	' 

			frm1.txtRel_cd.value = ""
			frm1.txtRel_nm.value = ""
			frm1.txtFamily_res_no.value = ""
			frm1.txtType.value = ""
			frm1.txtType_cd.value = ""
			frm1.txtEmp_no.focus

            txtFamily_Name_Onchange = true
        Else
			frm1.txtFamily_res_no.value = Trim(Replace(lgF0,Chr(11),""))
			frm1.txtRel_cd.value		= Trim(Replace(lgF1,Chr(11),""))
			frm1.txtRel_nm.value		= Trim(Replace(lgF2,Chr(11),""))
			frm1.txtType_cd.value		= Trim(Replace(lgF3,Chr(11),""))
			frm1.txtType.value			= Trim(Replace(lgF4,Chr(11),"")) 
        End if 
    End if  
    
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
		                            <TD CLASS="ctrow01">지급일자</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" ID="form01" NAME="txtMed_date" TYPE="Text" MAXLENGTH=10 SiZE=10 alt="지급일자" tag="22D" ondblclick="VBScript:Call OpenCalendar('txtMed_date',3)">
		                            </TD>
                                </TR>                               
                                  <TR>
		                            <TD CLASS="ctrow01">지급처상호</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtMed_name" ALT="지급처상호" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
		                            </TD>
                                </TR>  
                                <TR>
		                            <TD CLASS="ctrow01">지급처사업자번호</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtMed_rgst_no" ALT="지급처사업자번호" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22">
		                            '-' 제외</TD>
                                </TR> 
                                <TR>
		                            <TD CLASS="ctrow01">가족성명</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtFamily_Name" ALT="가족성명" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
										<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenCode(frm1.txtFamily_Name.value)">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>가족관계</TD>
		                            <TD CLASS="ctrow06"><INPUT TYPE=HIDDEN CLASS="form01" NAME="txtRel_cd" ALT="가족관계" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
														<INPUT CLASS="form01" NAME="txtRel_nm" ALT="가족관계" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01">주민번호</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtFamily_res_no" ALT="주민번호" TYPE="Text" MAXLENGTH=39 SiZE=20 tag="22">
		                            </TD>
                                </TR>                                 
                                <TR>
		                            <TD CLASS="ctrow01">대상자구분</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtType" ALT="대상자구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="22">
														<INPUT TYPE=HIDDEN NAME="txtType_cd" TAG="22">
		                            </TD>
                                </TR>
                                
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>제출구분</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtSubmit_cd" ALT="제출구분" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR> 
                                
                                  <TR>
		                            <TD CLASS="ctrow01" NOWRAP>지급건수</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtProvCnt" TYPE="Text" MAXLENGTH=13 SiZE=16 tag="22F" style='TEXT-ALIGN: right'>
		                            </TD>
                                </TR> 
                                
                                
		                        <TR>
		                            <TD CLASS="ctrow01">지급금액</TD>
		                            <TD CLASS="ctrow06"> <INPUT CLASS="form01" NAME="txtAmt" TYPE="Text" MAXLENGTH=13 SiZE=16 tag="22F" style='TEXT-ALIGN: right'>
		                            </TD>
                                </TR>
                                
                                
                                <TR>
		                            <TD CLASS="ctrow01">의료비내용</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtMed_text" ALT="의료비내용" TYPE="Text" MAXLENGTH=50 SiZE=50 tag="22">
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