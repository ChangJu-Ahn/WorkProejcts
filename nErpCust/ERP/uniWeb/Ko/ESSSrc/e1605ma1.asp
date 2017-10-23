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
    Dim dilig_dt
    Dim dilig_cd
	DIM dilig_hour,dilig_min
	dim emp_no
    dilig_dt = Trim(Request("dilig_dt"))
    dilig_cd = Trim(Request("Dilig_cd"))
    dilig_hour = Trim(Request("dilig_hour"))
    dilig_min = Trim(Request("dilig_min"))
	emp_no = Trim(Request("emp_no"))
%>

<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1605mb1.asp"						           '��: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim isOpenPop
Dim dilig_dt
Dim Dilig_cd
Dim gQuery_YN
dim fDiligAuth,fAuthCheck
Dim StartDate

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
    
	lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
    lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
    
    lgKeyStream = lgKeyStream & Trim(frm1.txtDilig_cd.value) & gColSep        
    lgKeyStream = lgKeyStream & UniConvDateAToB(Trim(frm1.txtFrDilig_dt.value),gDateFormat, gServerDateFormat) & gColSep
	if Trim(frm1.txtDilig_hour.value) = "" then
		lgKeyStream = lgKeyStream & "0" & gColSep 
	else 
		lgKeyStream = lgKeyStream & frm1.txtDilig_hour.value & gColSep  
	end if
	
	if Trim(frm1.txtDilig_min.value) = "" then
		lgKeyStream = lgKeyStream & "0" & gColSep        
	else 
		lgKeyStream = lgKeyStream & frm1.txtDilig_min.value & gColSep      
	end if      
    lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep          

	lgKeyStream = replace(lgKeyStream, "'", "''")

End Sub   
     
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx,i,ihour
    
    For i=0 To 23
    	Call SetCombo(frm1.txtFrDilig_hour, i, i)
    	Call SetCombo(frm1.txtToDilig_hour, i, i)
    	Call SetCombo(frm1.txtDilig_hour, i, i)
    Next
    
    ihour=0
    For i=0 To 5
    	Call SetCombo(frm1.txtFrDilig_min, ihour, ihour)
    	Call SetCombo(frm1.txtToDilig_min, ihour, ihour)
    	Call SetCombo(frm1.txtDilig_min, ihour, ihour)
    	ihour=ihour+10
    Next

	Call CommonQueryRs(" dilig_cd, dilig_nm "," hca010t ", " dilig_type=2 " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtdilig_cd, iCodeArr, iNameArr,Chr(11))
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
' Function Name : Form_Load
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '��: Clear err status
    dim dilig_hour,dilig_min,emp_no

	StartDate	= Date()
    dilig_dt = "<%=dilig_dt%>"
    Dilig_cd = "<%=dilig_cd%>"
	dilig_hour = "<%=dilig_hour%>"
	dilig_min = "<%=dilig_min%>"
	emp_no = "<%=emp_no%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode
    gQuery_YN = ""              
    call FncGetDiligAuth(fDiligAuth,fAuthCheck)

    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

	if emp_no  <> "" then
		 frm1.txtEmp_no.Value =emp_no  
	end if

    Call InitComboBox()
    Call LayerShowHide(0)

	Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
	if parent.txtName2.value = "" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if
    if  dilig_dt <> "" then
        frm1.txtFrDilig_dt.Value = dilig_dt
        frm1.txtFrDilig_hour.Value = dilig_hour
        frm1.txtFrDilig_min.Value = dilig_min
    else 
				frm1.txtFrDilig_dt.Value = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)
    end if
    if  dilig_dt <> "" then
        frm1.txtToDilig_dt.Value = dilig_dt
        frm1.txtToDilig_hour.Value = dilig_hour
        frm1.txtToDilig_min.Value = dilig_min
    else 
				frm1.txtToDilig_dt.Value = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)
    end if
    frm1.txtDilig_hour.Value = dilig_hour
    frm1.txtDilig_min.Value = dilig_min

    if  Dilig_cd <> "" then
        Call SetToolBar("01110")
        frm1.txtdilig_cd.value = Dilig_cd

        Call DbQuery(1)
    else    
        Call SetToolBar("01010")
        Call DbQueryEmp(1)
    end if
    Call LockField(Document)
    frm1.txtAppyn.value = "<%=Trim(Request("txtapp_yn"))%>"
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
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

    If Not chkFieldLength(Document, "1") Then									         '��: This function check required field
		Exit Function
	end if    
    Call LayerShowHide(1)
    
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                                               '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryEmp
'========================================================================================
Function DbQueryEmp(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '��: Clear err status
    DbQueryEmp = False                                                              '��: Processing is NG
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQueryEmp = True                                                               '��: Processing is NG
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
		ProtectTag(frm1.txtDilig_cd)
		ProtectTag(frm1.txtFrDilig_dt)
		
		if frm1.txtAppyn.value = "Y" or frm1.txtAppyn.value = "R" then
			ProtectTag(frm1.txtDilig_hour)
			ProtectTag(frm1.txtDilig_min)		
			ProtectTag(frm1.txtRemark)
			ProtectTag(frm1.txtApp_emp_no)
			frm1.btnCalType.disabled = true
			Call SetToolBar("01000")
		else
			Call SetToolBar("01110")
		end if
		
    End if 
	
    frm1.txtDilig_cd.disabled = true
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
	dim datechk
	On Error Resume Next
    Err.Clear                                                                    '��: Clear err status
	
	DbSave = False	
	'----�ݷ�ó���� ���׿� ���ؼ��� ������ �̷���� �� ����.
	strAppyn = frm1.txtAppyn.value 
	If strAppyn = "R" Then
		Call DisplayMsgBox("800477","X","X","X")
        exit function
	End IF
	
	With Frm1
        if .txtEmp_no.value = "" then
            Call DisplayMsgBox("800006","X","X","X")
            .txtDilig_cd.focus()
            exit function
        end if
        if  Date_chk(.txtFrDilig_dt.value, strDate) = True then
            .txtFrDilig_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtFrDilig_dt.focus()
            exit function
        end if

        if .txtDilig_hour.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_hour.focus()
            exit function
        end if
		if Trim(.txtDilig_min.value)="" then
			.txtDilig_min.value = 0
		end if        
        if .txtDilig_cd.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_cd.focus()
            exit function
        end if        
        if .txtApp_emp_no.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtApp_emp_no.focus()
            exit function
        end if		
	 '----����ڿ� �����ڰ� ������ ����.
        if .txtEmp_no.value = .txtApp_emp_no.value  then 
            Call DisplayMsgBox("800476","X","X","X")
		        .txtapp_emp_no.value    = ""
		        .txtapp_name.value      = ""
		        .txtApp_emp_no.focus()
            exit function
        end if
	'-------�����ڰ� �����ϴ� ������� 
	if txtApp_emp_no_Onchange() = False then
		Exit Function
	end if
	
	End With    
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
' Function Name :  FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	lgIntFlgMode = OPMD_CMODE              'create mode

    Call ClearField(document,2)
    Call LockField(Document)	
    Call SetToolBar("01010")

	frm1.txtFrDilig_dt.value = UniConvDateAToB(Date(),gServerDateFormat,gDateFormat)

    frm1.txtDilig_cd.focus()    
    frm1.txtAppyn.value = ""
    frm1.txtApp_emp_no.value = ""

    FncNew = True																 '��: Processing is OK
End Function

'========================================================================================================
' Function Name :  GetRow
'========================================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================================
'   Event Name : txtApp_emp_no_Onchange()            '<==������ �̸��������� 
'========================================================================================================
Function txtApp_emp_no_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no
    
    IF frm1.txtApp_emp_no.value = "" THEN
        frm1.txtApp_name.value = ""
        txtApp_emp_no_Onchange = true
    ELSE
		strEmp_no = frm1.txtApp_emp_no.value 

        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(frm1.txtApp_emp_no.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
'			Call DisplayMsgbox("800048","X","X","X")	'���޳����ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
			Call DisplayMsgBox("970000","X","���","X")

            frm1.txtApp_name.value = ""
            frm1.txtApp_emp_no.focus
        ELSE    
            frm1.txtApp_name.value = Trim(Replace(lgF0,Chr(11),""))   '�����ڵ� 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
End Function

'========================================================================================================
' Name : OpenEmp()
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)
	If IsOpenPop = True or  frm1.txtAppyn.value = "Y"  or  frm1.txtAppyn.value = "R" Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtApp_Emp_no.value			' Code Condition
	arrParam(1) = ""								' Name Cindition
    arrParam(2) = "APPROVAL_CODE"					'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa3.asp", Array(arrParam), _
		"dialogWidth=546px; dialogHeight=387px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
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

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Function txtApp_emp_no_onKeyDown()
	Dim CuEvObj,KeyCode,IntRetCd
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	txtApp_emp_no_onKeyDown	= True	
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
            frm1.txtApp_name.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '����� �ش��ϴ� �̸� 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
End Function

Sub txtDilig_cd_OnChange()
    gQuery_YN = "Y"                                                    ' ������ ü������ �ν��Ѵ�.�������� 
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function

Function CalHHMM()
	Dim iFrDate, iFrHour, iFrMinute
	Dim iToDate, iToHour, iToMinute
	Dim iWorkingMinutes
	
	If frm1.txtFrDilig_dt.value="" Then Exit Function
	If frm1.txtFrDilig_hour.value="" Then Exit Function
	If frm1.txtFrDilig_min.value="" Then Exit Function
	If frm1.txtToDilig_dt.value="" Then Exit Function
	If frm1.txtToDilig_hour.value="" Then Exit Function
	If frm1.txtToDilig_min.value="" Then Exit Function

	iFrDate = frm1.txtFrDilig_dt.value
	iFrHour = right("00" & frm1.txtFrDilig_hour.value, 2)
	iFrMinute = right("00" & frm1.txtFrDilig_min.value, 2)
	iFrDate = iFrDate & " " & iFrHour & ":" & iFrMinute & ":00"
	If isDate(iFrDate) = False Then Exit Function
		
	iToDate = frm1.txtToDilig_dt.value
	iToHour = right("00" & frm1.txtToDilig_hour.value, 2)
	iToMinute = right("00" & frm1.txtToDilig_min.value, 2)
	iToDate = iToDate & " " & iToHour & ":" & iToMinute & ":00"
	If isDate(iToDate) = False Then Exit Function

	frm1.txtDilig_STRT_dt.value = iFrDate
	frm1.txtDilig_END_dt.value = iToDate
	
	iWorkingMinutes = datediff("n",iFrDate,iToDate)
	
	
	frm1.txtDilig_hour.value = Int(iWorkingMinutes / 60)
	frm1.txtDilig_min.value = iWorkingMinutes mod 60
	
	
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
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">���</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">����</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
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
		                            <TD CLASS="ctrow01">�ٹ�����</TD>
		                            <TD CLASS="ctrow06" COLSPAN=3><SELECT  CLASS="form01" NAME="txtDilig_cd" ALT="�ٹ�����" STYLE="WIDTH: 150px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01">�ٹ��Ⱓ(FROM)</TD>
		                            <TD CLASS="ctrow06" COLSPAN=3>
										<INPUT CLASS="form01" ID="txtFrDilig_dt" NAME="txtFrDilig_dt" TYPE="Text" MAXLENGTH=10 SiZE=12 alt="�ٹ���¥" STYLE="text-align: center" TAG="22D" onChange="vbscript:CalHHMM()" ondblclick="VBScript:Call OpenCalendar('txtFrDilig_dt',3)">&nbsp;
										<SELECT CLASS="form01" NAME="txtFrDilig_hour" ALT="��" STYLE="WIDTH: 50px" TAG="22" onChange="vbscript:CalHHMM()"><OPTION VALUE=""></OPTION></SELECT>��
										<SELECT CLASS="form01" NAME="txtFrDilig_min" ALT="��" STYLE="WIDTH: 50px" TAG="22" onChange="vbscript:CalHHMM()"><OPTION VALUE=""></OPTION></SELECT>��
&nbsp;~&nbsp;
										<INPUT CLASS="form01" ID="txtToDilig_dt" NAME="txtToDilig_dt" TYPE="Text" MAXLENGTH=10 SiZE=12 alt="�ٹ���¥" STYLE="text-align: center" TAG="22D" onChange="vbscript:CalHHMM()" ondblclick="VBScript:Call OpenCalendar('txtFrDilig_dt',3)">&nbsp;
										<SELECT CLASS="form01" NAME="txtToDilig_hour" ALT="��" STYLE="WIDTH: 50px" TAG="22" onChange="vbscript:CalHHMM()"><OPTION VALUE=""></OPTION></SELECT>��
										<SELECT CLASS="form01" NAME="txtToDilig_min" ALT="��" STYLE="WIDTH: 50px" TAG="22" onChange="vbscript:CalHHMM()"><OPTION VALUE=""></OPTION></SELECT>��
		                            </TD>
                                </TR>  
                                <TR>
		                            <TD CLASS="ctrow01">���½ð�</TD>
		                            <TD CLASS="ctrow06" COLSPAN=3>
										<SELECT CLASS="form01" NAME="txtDilig_hour" ALT="�ð�" STYLE="WIDTH: 50px" TAG="24"><OPTION VALUE=""></OPTION></SELECT>�ð�
										<SELECT CLASS="form01" NAME="txtDilig_min" ALT="��" STYLE="WIDTH: 50px" TAG="24"><OPTION VALUE=""></OPTION></SELECT>��
		                            </TD>
                                </TR>                                
                                <TR>
		                            <TD CLASS="ctrow01">����</TD>
		                            <TD CLASS="ctrow06" COLSPAN=3><INPUT CLASS="form01" NAME="txtRemark" ALT="����" TYPE="Text" MAXLENGTH=39 SiZE=40 TAG="22">
		                            </TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01">������</TD>
		                            <TD CLASS="ctrow06" colspan=3><INPUT CLASS="form01" NAME="txtApp_emp_no" ALT="���λ��" TYPE="Text" MAXLENGTH=13 SiZE=13 TAG="22">&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtApp_emp_no.value)">
		                                    <INPUT CLASS="form02" NAME="txtApp_name" TYPE="Text" MAXLENGTH=20 SiZE=20 TAG="24">
		                            </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0  border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtAppyn"		 TAG="24">    
    <INPUT TYPE=HIDDEN NAME="txtDilig_STRT_dt"		 TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtDilig_END_dt"		 TAG="24">    

</FORM>	

</BODY>
</HTML>
