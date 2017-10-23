<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<HTML>
<HEAD>
<TITLE><%=Request("title")%></TITLE>

<!-- #Include file="../ESSinc/IncServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->

<Script Language="VBScript">

Option Explicit 

Const BIZ_PGM_ID      = "e1101mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim isOpenPop

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
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
End Sub        
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
'   결혼구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0105", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtmarry_cd, iCodeArr, iNameArr, Chr(11))    

	' 혈액형1
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0106", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtBlood_type1, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0107", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtBlood_type2, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0015", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txthouse_cd, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0028", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtmemo_cd, iCodeArr, iNameArr, Chr(11))    

End Sub
'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ClearField(document,2)
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    if  gDeptAuth = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        Call SetToolBar("10010")
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

    Call InitComboBox()
    Call LayerShowHide(0)
    Call LockField(Document)
   
    Call DbQuery(1)
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

    Call MakeKeyStream("Q")

    Call LayerShowHide(1)

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
	lgIntFlgMode = OPMD_UMODE              'update mode

    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        if  frm1.txtEmp_no.value = parent.txtEmp_no.Value then
            Call SetToolBar("10010")
        else
            Call SetToolBar("10000")
        end if
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	Dim strWhere
	dim ok
	dim temp
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG

    if  frm1.txtso_lu_cd1.checked = true then
        frm1.txtso_lu_cdv.value = "1"
    else
        frm1.txtso_lu_cdv.value = "2"
    end if

    if  not(frm1.txtCurr_addr.value = "") and frm1.txtCurr_zip_cd.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            frm1.txtCurr_zip_cd.focus()
            exit function
    end if

    if  frm1.txtHgt.value = "" then
		frm1.txtHgt.value = 0
    Else
		 if num_chk(frm1.txtHgt.value, temp) = false then
            Call DisplayMsgBox("229924","X","X","X")
            frm1.txtHgt.focus()
            exit function
         elseif temp < 0 then
			Call DisplayMsgBox("800484","X",frm1.txtHgt.alt ,"X")
            frm1.txtHgt.focus()
            exit function
        end if
    end if

	
    if  frm1.txtWgt.value = "" then
		frm1.txtWgt.value = 0
    Else
        if  num_chk(frm1.txtWgt.value,temp)= false then
            Call DisplayMsgBox("229924","X","X","X")
            frm1.txtWgt.focus()
            exit function
        elseif temp < 0 then
			Call DisplayMsgBox("800484","X",frm1.txtWgt.alt ,"X")
            frm1.txtWgt.focus()
            exit function
        end if
    end if
	
	
    if  frm1.txtEyesgt_left.value = "" then
		frm1.txtEyesgt_left.value = 0
    Else
        if  Num_chk(frm1.txtEyesgt_left.value, temp) = True then
			if abs(temp) >= 10 then
				Call DisplayMsgBox("800094","X","X","X")
				frm1.txtEyesgt_left.focus()
				exit function
			end if
        else
            Call DisplayMsgBox("229924","X","X","X")
            frm1.txtEyesgt_left.focus()
            exit function
        end if
    end if
	
    if  frm1.txtEyesgt_right.value = "" then
		frm1.txtEyesgt_right.value = 0
    Else
		if  Num_chk(frm1.txtEyesgt_right.value, temp) = True then
			if abs(temp) >= 10 then
				Call DisplayMsgBox("800094","X","X","X")
				frm1.txtEyesgt_right.focus()
				exit function
			end if
		else
			Call DisplayMsgBox("229924","X","X","X")
			frm1.txtEyesgt_right.focus()
			exit function
		end if
    end if

    if frm1.txtbirt.value = "" then
    elseif Date_chk(frm1.txtbirt.value, strDate) = True then
        frm1.txtbirt.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtbirt.focus()
        exit function
    end if	

    if frm1.txtMemo_cd.value = "" and frm1.txtMemo_dt.value <>"" then
		Call DisplayMsgBox("800094","X","X","X")
		frm1.txtMemo_cd.focus()
	    exit function
    else
		if frm1.txtMemo_dt.value = "" then    
		elseif  Date_chk(frm1.txtMemo_dt.value, strDate) = True then
 
		     frm1.txtMemo_dt.value = strDate
		else
		    Call DisplayMsgBox("800094","X","X","X")
		    frm1.txtMemo_dt.focus()
		    exit function
		end if
	end if
'우편번호 체크 

    If frm1.txtZip_cd.value <> "" then
        strWhere =                " ZIP_CD =  " & FilterVar(frm1.txtZip_cd.value , "''", "S") & ""
        strWhere = strWhere & " AND COUNTRY_CD=  " & FilterVar(frm1.txtNat_cd.value , "''", "S") & ""
        if  CommonQueryRs(" COUNT(*) "," B_ZIP_CODE ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
            if  Replace(lgF0, Chr(11), "") = 0 then
		          Call DisplayMsgBox("800016","X","X","X")
                  frm1.txtZip_cd.focus()
                  exit function
            end if
        end if
    End If
    
    If frm1.txtCurr_zip_cd.value <> "" then
        strWhere =                " ZIP_CD = " & FilterVar(frm1.txtCurr_zip_cd.value , "''", "S") & ""
        strWhere = strWhere & " AND COUNTRY_CD= " & FilterVar(frm1.txtNat_cd.value , "''", "S") & ""

        if  CommonQueryRs(" COUNT(*) "," B_ZIP_CODE ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
            if  Replace(lgF0, Chr(11), "") <= 0 then
		          Call DisplayMsgBox("800016","X","X","X")
                  frm1.txtCurr_zip_cd.focus()
                  exit function
            end if
        end if
    End If
 
    If Not chkFieldLength(Document, "2") Then									         '☜: This function check required field
		Exit Function
	end if

	Call LayerShowHide(1)
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
    Call DbQuery(1)
End Function

'========================================================================================================
' Name : OpenZip()
'========================================================================================================

Function OpenZip(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere	
		Case 1
	        arrParam(0) = frm1.txtZip_cd.value
	        arrParam(1) = ""
            arrParam(2) = frm1.txtNat_cd.value
		Case 2
	        arrParam(0) = frm1.txtCurr_zip_cd.value
	        arrParam(1) = ""
            arrParam(2) = frm1.txtNat_cd.value
	End Select    

	arrRet = window.showModalDialog("E1ZipPopa1.asp", Array(arrParam), _
		"dialogWidth=466px; dialogHeight=366px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	End If	

	Select Case iWhere	
		Case 1
	        frm1.txtZip_cd.value = arrRet(0)
	        frm1.txtAddr.value = arrRet(1)
	        frm1.txtAddr.focus()
		Case 2
	        frm1.txtCurr_Zip_cd.value = arrRet(0)
	        frm1.txtCurr_Addr.value = arrRet(1)
	        frm1.txtCurr_Addr.focus()
	End Select    
			
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

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 target=MyBizASP METHOD="POST">
<FORM NAME=frm1 target=MyBizASP METHOD="POST">
    <TABLE width=715 cellSpacing=0 cellPadding=0 border=0>
        <TR>
           <td height="10"></td>
        </TR>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=1 cellPadding=0 border=0 bgcolor=#DDDDDD>
                    <TR>
                        <TD CLASS="ctrow01">사원</TD>
                        <TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtEmp_no" ALT="사번" TYPE="Text" MAXLENGTH=13 SiZE=13 readonly>
                            <INPUT CLASS="form02" NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=15 SiZE=16 readonly>
                        </TD>
	            		<TD CLASS="ctrow01">직위</TD>
	            		<TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtroll_pstn" ALT="직위" TYPE="Text" MAXLENGTH=20 SiZE=20 readonly>
                        </TD>
                    <TR>
	            		<TD CLASS="ctrow01">부서</TD>
	            		<TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtDept_nm" ALT="부서" TYPE="Text" MAXLENGTH=20 SiZE=20 readonly>
                        </TD>
	            		<TD CLASS="ctrow01">최근승급일</TD>
	            		<TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtresent_promote_dt" ALT="최근승급일" TYPE="Text" MAXLENGTH=10 SiZE=14 readonly>
                        </TD>
                    </TR>

                    <TR>
                        <TD CLASS="ctrow01">그룹입사일</TD>
                        <TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtGroup_entr_dt" ALT="그룹입사일" TYPE="Text" MAXLENGTH=10 SiZE=14 readonly>
                        </TD>
	            		<TD CLASS="ctrow01">입사일</TD>
	            		<TD CLASS="ctrow02">
                            <INPUT CLASS="form02" NAME="txtEntr_dt" ALT="입사일" TYPE="Text" MAXLENGTH=10 SiZE=14 readonly>
                        </TD>
                    </TR>

                    <TR>
	            		<TD CLASS="ctrow01">영문성명</TD>
	            		<TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txteng_name" ALT="영문성명" TYPE="Text" MAXLENGTH=30 SiZE=30>
                        </TD>
                    </TR>
                    <TR>
                        <TD CLASS="ctrow01">생년월일</TD>
                        <TD CLASS="ctrow02">
                            <INPUT CLASS="form01" ID="txtbirt" NAME="txtbirt" ALT="생년월일" TYPE="Text" MAXLENGTH=10 SiZE=14 ondblclick="VBScript:Call OpenCalendar('txtbirt',3)">&nbsp;&nbsp;
                           	<INPUT TYPE="RADIO" CLASS="ftgray" NAME="txtso_lu_cd" CHECKED ID="txtso_lu_cd1" VALUE=1><LABEL FOR="txtso_lu_cd1">양력</LABEL>
    					    <INPUT TYPE="RADIO" CLASS="ftgray" NAME="txtso_lu_cd" ID="txtso_lu_cd2" VALUE=2><LABEL FOR="txtso_lu_cd2">음력</LABEL>
    					    <INPUT CLASS="form01" TYPE="hidden" name="txtso_lu_cdv"  ID="txtso_lu_cdv">
                        </TD>
	            		<TD CLASS="ctrow01">기념일</TD>
	            		<TD CLASS="ctrow02">
	            		    <SELECT CLASS="form01" NAME="txtMemo_cd" CLASS="ftgray" ALT="기념일구분" STYLE="WIDTH: 100px"><OPTION VALUE=""></OPTION></SELECT>
                            <INPUT CLASS="form01" ID="txtMemo_dt" NAME="txtMemo_dt" ALT="기념일" TYPE="Text" MAXLENGTH=10 SiZE=14 ondblclick="VBScript:Call OpenCalendar('txtMemo_dt',3)">
                        </TD>
                    </TR>
                    <TR>
	            		<TD CLASS=ctrow01>결혼여부</TD>
	            		<TD CLASS=ctrow02 align=left>
	            		    <SELECT NAME="txtMarry_Cd" CLASS=form01 ALT="결혼여부" STYLE="WIDTH: 100px"><OPTION VALUE=""></OPTION></SELECT>
	            		</TD>
	            		<TD CLASS="ctrow01">주거구분</TD>
	            		<TD CLASS="ctrow02">
	            		    <SELECT NAME="txtHouse_Cd" CLASS=form01 ALT="주거구분" STYLE="WIDTH: 100px"><OPTION VALUE=""></OPTION></SELECT>
                        </TD>
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow01">본적</TD>
	                    <TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txtDomi" ALT="본적" TYPE="Text" MAXLENGTH=80 SiZE=80>
                        </TD>
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow01" rowspan=2>주민등록지</TD>
	                    <TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txtZip_cd" ALT="주민등록지우편번호" TYPE="Text" MAXLENGTH=13 SiZE=13>&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZip(1)">
                        </TD>      
                    </TR>
                    <TR>

	                    <TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txtAddr" ALT="주민등록지" TYPE="Text" MAXLENGTH=128 SiZE=80>
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow01" rowspan=2>현주소</TD>
	                    <TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txtCurr_zip_cd" ALT="현주소우편번호" TYPE="Text" MAXLENGTH=13 SiZE=13>&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZip(2)">
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow02" colspan=3>
                            <INPUT CLASS="form01" NAME="txtCurr_addr" ALT="현주소" TYPE="Text" MAXLENGTH=128 SiZE=80>
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow01">전화번호</TD>
	                    <TD CLASS="ctrow02"><INPUT CLASS="form01" NAME="txtTel_no" ALT="전화번호" TYPE="Text" MAXLENGTH=20 SiZE=20></TD>
	                    <TD CLASS="ctrow01">비상연락번호</TD>
	                    <TD CLASS="ctrow02"><INPUT CLASS="form01" NAME="txtEm_tel_no" ALT="비상연락번호" TYPE="Text" MAXLENGTH=20 SiZE=20></TD>
                    </TR>
                    <TR>
	                    <TD CLASS="ctrow01">E-Mail</TD>
	                    <TD CLASS="ctrow02">
                            <INPUT CLASS="form01" NAME="txtEmail_addr" ALT="E-Mail" TYPE="Text" MAXLENGTH=30 SiZE=30>
                        </TD>      
	                    <TD CLASS="ctrow01">핸드폰</TD>
	                    <TD CLASS="ctrow02">
                            <INPUT CLASS="form01" NAME="txtHand_tel_no" ALT="핸드폰" TYPE="Text" MAXLENGTH=20 SiZE=20>
                        </TD>      
                    </TR>
                    <TR>
	            	    <TD CLASS="ctrow01">신장</TD>
	            	    <TD CLASS="ctrow02"><INPUT CLASS="form01" NAME="txtHgt" ALT="신장" TYPE="Text" MAXLENGTH=5 SiZE=5>&nbsp;Cm</TD>
	            	    <TD CLASS="ctrow01">체중</TD>
	            	    <TD CLASS="ctrow02"><INPUT CLASS="form01" NAME="txtWgt" ALT="체중" TYPE="Text" MAXLENGTH=5 SiZE=5>&nbsp;Kg</TD>
                    </TR>
                    <TR>
	            	    <TD CLASS="ctrow01">혈액형</TD>
	            	    <TD CLASS="ctrow02"><SELECT NAME="txtBlood_type1" CLASS=form01 ALT="혈액형1" STYLE="WIDTH: 70px"><OPTION VALUE=""></OPTION></SELECT>&nbsp;형
	            	                        <SELECT NAME="txtBlood_type2" CLASS=form01 ALT="혈액형2" STYLE="WIDTH: 70px"><OPTION VALUE=""></OPTION></SELECT></TD>
	            	    <TD CLASS="ctrow01">색맹</TD>
	            	    <TD CLASS="ctrow02"><INPUT CLASS="frgray" TYPE=CHECKBOX NAME="txtDalt_type"></TD>
                    </TR>

                    <TR>
	            	    <TD CLASS="ctrow01">시력</TD>
	            	    <TD CLASS="ctrow02">&nbsp;좌<INPUT CLASS="form01" NAME="txtEyesgt_left" ALT="시력(좌)" TYPE="Text" MAXLENGTH=4 SiZE=5>
                                             &nbsp;우<INPUT CLASS="form01" NAME="txtEyesgt_right" ALT="시력(우)" TYPE="Text" MAXLENGTH=4 SiZE=5></TD>
                        <TD CLASS="ctrow01"></TD>
                        <TD CLASS="ctrow02"></TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=5></TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtNat_cd">

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId"> 
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">    
</FORM>	
</BODY>
</HTML>
