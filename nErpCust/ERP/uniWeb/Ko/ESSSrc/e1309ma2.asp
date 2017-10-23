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
    Dim strYear, strEmp_no, strFamilyName,strRel_cd,strType,strAmt,strMainInsertFlag
    
    strYear = TRIM(Request("txtYear"))
    strEmp_no  = TRIM(Request("txtEmp_no"))
    strFamilyName      = TRIM(Request("txtFamilyName"))
    strRel_cd      = TRIM(Request("txtRel_cd"))
    strType      = TRIM(Request("txtType"))
    strAmt      = TRIM(Request("txtAmt"))   
    strMainInsertFlag  = TRIM(Request("txtMainInsertFlag")) 
%>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "e1309mb2.asp"						           '��: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim isOpenPop
Dim strYear, strEmp_no, strFamilyName,strRel_cd,strType,strAmt,strMainInsertFlag

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
	lgKeyStream = lgKeyStream & Trim(frm1.txtFamilyName.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtRel_cd.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtType_cd.value) & gColSep 
	lgKeyStream = lgKeyStream & Trim(frm1.txtAmt.value) & gColSep  
	lgKeyStream = lgKeyStream & Trim(frm1.txtSubmit_Cd.value) & gColSep      

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
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0024'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.txtType_cd, iCodeArr, iNameArr, Chr(11))
    
   
 
     iCodeArr="Y" & Chr(11) &"N" & Chr(11)
     iNameArr="����û�ڷ�" & Chr(11) &"�׹����ڷ�" & Chr(11)
     
    Call SetCombo2(frm1.txtSubmit_Cd, iCodeArr, iNameArr, Chr(11))  
        
    
   
    
        
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = 4+1
    Grid1.SheetMaxrows = 3
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '��: Clear err status

 
    strYear = "<%=strYear%>"
    If "<%=strEmp_no%>" <> "" Then
		strEmp_no = "<%=strEmp_no%>"
	Else
		strEmp_no = parent.txtEmp_no2.Value
	End IF
	strFamilyName  = "<%=strFamilyName%>"  
    strRel_cd = "<%=strRel_cd%>"
    strType = "<%=strType%>"
    strAmt = "<%=strAmt%>"
    strMainInsertFlag = "<%=strMainInsertFlag%>"
	ProtectTag(frm1.txtRel_nm)	    
	'

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

	if  strFamilyName <> "" then
	    Call SetToolBar("01110")

	    frm1.txtYear.value			= strYear
	    frm1.txtFamilyName.value	= strFamilyName
	    frm1.txtType_cd.value		= strType
	   frm1.txtSubmit_cd.value		= "<%=Request("txtflag")%>" 
	    Call DbQuery(1)
	else    
	    Call SetToolBar("01010")

	    Call DbQuery(1)        
	end if
	frm1.txtSubmit_Cd.value ="N"
	ProtectTag(frm1.txtType_cd)
End Sub

'========================================================================================
' Function Name : Window_onUnLoad
'========================================================================================
Sub Form_unLoad()
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    Dim iret
    Err.Clear                                                                    '��: Clear err status
	
    DbQuery = False                                                              '��: Processing is NG

    If Not chkFieldLength(Document, "1") Then									         '��: This function check required field
		Exit Function
	end if    
    Call LayerShowHide(1)
 
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
    strVal = strVal     & "&strMainInsertFlag="       & strMainInsertFlag
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                                               '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '��: Clear err status
 
    If gQuery_YN = "Y" Then                    '�����ڵ� ������ ü������ �Ͼ���.
		lgIntFlgMode = OPMD_CMODE              'create mode
    Else                                       '���ϱ��� ��Ȳ���� update�� �Ϸ��� ��� ������ 
		lgIntFlgMode = OPMD_UMODE              'update mode
		ProtectTag(frm1.txtYear)
		ProtectTag(frm1.txtFamilyName)
		ProtectTag(frm1.txtRel_nm)
		ProtectTag(frm1.txtType_cd)
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
    Err.Clear                                                                    '��: Clear err status

	DbSave = False														         '��: Processing is NG
    
    if  frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtYear.focus()
        exit function
    end if
 
    if  frm1.txtFamilyName.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtFamilyName.focus()
        exit function
    end if

    'If txtFamilyName_Onchange() Then      
        'Exit Function
  '  End if
    
    if  frm1.txtType_cd.value = "" then
        'Call DisplayMsgBox("800094","X","X","X")
        'frm1.txtType_cd.focus()
        'exit function
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
                
	If Not chkFieldLength(Document, "2") Then									         '��: This function check required field
       Exit Function
    End If
    
    Call MakeKeyStream("C")

	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '��: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave  = True                                                               '��: Processing is NG

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
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("C")
	With Frm1
		.txtMode.value        = "UID_M0003"                                        '��: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                              '��: Processing is NG
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

    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	lgIntFlgMode = OPMD_CMODE              'create mode

    Call ClearField(document,2)
	
	frm1.txtYear.selectedIndex = 0

    Call LockField(Document)	
	Call ProtectTag(frm1.txtRel_nm)    
    Call SetToolBar("01010")
    frm1.txtSubmit_cd.value ="N"

    FncNew = True																 '��: Processing is OK

End Function


'========================================================================================================
' Name : FncGetDiligAuth()
' Desc : developer describe this line 
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
Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function


'========================================================================================================
' Name : OpenCode()
'========================================================================================================
Function OpenCode(pCode)
	Dim arrRet
	Dim arrParam(4)
    dim IntRetCd
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HFA150T"  								' TABLE
	arrParam(1) = "family_name"								' Code Condition
	arrParam(2) = "dbo.ufn_GetCodeName('H0140',FAMILY_REL) "								' Name Cindition
	arrParam(3) =  frm1.txtFamilyName.value					' Code�� 
	arrParam(4) = " INSUR_YN ='Y' and year_yy = " & FilterVar(frm1.txtYear.value, "''", "S") & " and emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " and family_name like " 	'WHERE���� 

	arrRet = window.showModalDialog("e1codepopa1.asp", Array(arrParam), _
		"dialogWidth=546px; dialogHeight=340px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	
		
	If arrRet(0) = "" Then
		Exit Function
	Else
	
		frm1.txtType_cd.value = ""
 		frm1.txtFamilyName.value = arrRet(0)
		frm1.txtRel_nm.value = arrRet(1)
		
	
	 		IntRetCd = CommonQueryRs(" top 1 MINOR_CD,MINOR_NM , PARIA_YN "," HFA150T  left outer join  B_MINOR   on HFA150T.FAMILY_REL= B_MINOR.MINOR_CD ",_
 		 " MAJOR_CD ='H0140' AND INSUR_YN = 'Y' AND FAMILY_NAME = " & FilterVar(frm1.txtFamilyName.value, "''", "S")  &_
 		 " AND EMP_NO= " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " AND YEAR_YY = " & FilterVar(frm1.txtYear.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 		 
		frm1.txtRel_cd.value = Trim(Replace(lgF0,Chr(11),""))
	
		call CommonQueryRs(" supp_cd, dbo.ufn_GetCodeName('H0024',supp_cd) ,rel_cd", " HAA020T ",_
				" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND FAMILY_NM= " & FilterVar( arrRet(0), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	

		if Trim(Replace(lgF0,Chr(11),""))="" then
		
		else	
			frm1.txtType_cd.value = Trim(Replace(lgF0,Chr(11),""))
		       	       
		end if	
		       	

 	End If	
		
End Function
 '========================================================================================================
'   Event Name :txtFamily_Name_Onchange    
'========================================================================================================
Function txtFamilyName_Onchange()
    Dim IntRetCd
    gQuery_YN = "Y"      
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtRel_cd.value = ""
		frm1.txtRel_nm.value = ""
		frm1.txtType_cd.value = ""
    Else

 		IntRetCd = CommonQueryRs(" top 1 MINOR_CD,MINOR_NM , PARIA_YN "," HFA150T  left outer join  B_MINOR   on HFA150T.FAMILY_REL= B_MINOR.MINOR_CD ",_
 		 " MAJOR_CD ='H0140' AND INSUR_YN = 'Y' AND FAMILY_NAME = " & FilterVar(frm1.txtFamilyName.value, "''", "S")  &_
 		 " AND EMP_NO= " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " AND YEAR_YY = " & FilterVar(frm1.txtYear.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	    if  IntRetCd = False then
   			Call  DisplayMsgBox("971012","X","�ξ簡�������� ������ �׸� üũ�� �Ǿ� �ִ���","X")	' 

			frm1.txtRel_cd.value = ""
			frm1.txtRel_nm.value = ""
			frm1.txtType_cd.value = ""
			frm1.txtEmp_no.focus

            txtFamilyName_Onchange = true
        Else
			frm1.txtRel_cd.value		= Trim(Replace(lgF0,Chr(11),""))
			frm1.txtRel_nm.value		= Trim(Replace(lgF1,Chr(11),""))

					
        End if 
        
        call CommonQueryRs(" supp_cd, dbo.ufn_GetCodeName('H0024',supp_cd) ,rel_cd", " HAA020T ",_
				" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND FAMILY_NM= " & FilterVar( txtFamilyName, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	

		if Trim(Replace(lgF0,Chr(11),""))="" then
		
		else	
			frm1.txtType_cd.value = Trim(Replace(lgF0,Chr(11),""))
		       	       
		end if	
		       	
		       	
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
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">���</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">�μ�</td>
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
		                            <TD CLASS="ctrow01" NOWRAP>���꿬��</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtYear" ALT="���꿬��" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR> 
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>��������</TD>
		                            <TD CLASS="ctrow06"><INPUT CLASS="form01" NAME="txtFamilyName" ALT="��������" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
										<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenCode(frm1.txtFamilyName.value)">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>��������</TD>
		                            <TD CLASS="ctrow06"><INPUT TYPE=HIDDEN CLASS="form01" NAME="txtRel_cd" ALT="��������" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
														<INPUT CLASS="form01" NAME="txtRel_nm" ALT="��������" TYPE="Text" MAXLENGTH=39 SiZE=30 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>����</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtType_cd" ALT="����" STYLE="WIDTH: 150px" TAG="24"><OPTION VALUE="">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>���ⱸ��</TD>
		                            <TD CLASS="ctrow06"><SELECT CLASS="form01" NAME="txtSubmit_cd" ALT="���ⱸ��" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR> 
                                
                                
                                
		                        <TR>
		                            <TD CLASS="ctrow01" NOWRAP>�ݾ�</TD>
		                            <TD CLASS="ctrow06"> <INPUT class="form01" NAME="txtAmt" TYPE="Text" MAXLENGTH=13 SiZE=16 tag="22" style='TEXT-ALIGN: right'>
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

    <INPUT TYPE=HIDDEN NAME="txtAppyn"    TAG="24">
</FORM>	

</BODY>
</HTML>