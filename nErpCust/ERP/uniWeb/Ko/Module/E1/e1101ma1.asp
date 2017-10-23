<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("title")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncServer.asp"  -->

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

Const BIZ_PGM_ID      = "e1101mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim isOpenPop


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
        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        else
            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
        end if
    else        
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
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

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ClearField(document,2)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    FncNew = True																 '☜: Processing is OK
End Function
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
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
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing

End Sub

Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    Call MakeKeyStream("Q")

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

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

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	Dim strWhere
	dim ok
	dim temp
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
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
        strWhere =                " ZIP_CD =  " & FilterVar(frm1.txtCurr_zip_cd.value , "''", "S") & ""
        strWhere = strWhere & " AND COUNTRY_CD=  " & FilterVar(frm1.txtNat_cd.value , "''", "S") & ""

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

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG


End Function



'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)
End Function




Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'===================================================================
'========================================================================================================
' Name : OpenZip()
' Desc : developer describe this line 
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
		"dialogWidth=470px; dialogHeight=385px; center: Yes; help: No; resizable: No; status: No;")
		
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
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub
Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
</SCRIPT>

<!-- #Include file="../../inc/uniSimsClassID.inc" --> 

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 target=MyBizASP METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=749>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=1 cellPadding=0 border=0 width=721 bgcolor=#ffffff>
                    <TR>
                        <TD CLASS="TDFAMILY_TITLE" NOWRAP>사원</TD>
                        <TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEmp_no" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="24">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtName" TYPE="Text" MAXLENGTH=15 SiZE=10 tag="24">
                        </TD>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>직위</TD>
	            		<TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtroll_pstn" TYPE="Text" MAXLENGTH=20 SiZE=10 tag="24">
                        </TD>
                    <TR>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>부서</TD>
	            		<TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDept_nm" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
                        </TD>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>최근승급일</TD>
	            		<TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtresent_promote_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24">
                        </TD>
                    </TR>

                    <TR>
                        <TD CLASS="TDFAMILY_TITLE" NOWRAP>그룹입사일</TD>
                        <TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtGroup_entr_dt" ALT="그룹입사일" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24">
                        </TD>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>입사일</TD>
	            		<TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEntr_dt" ALT="입사일" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24">
                        </TD>
                    </TR>

                    <TR>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>영문성명</TD>
	            		<TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txteng_name" ALT="영문성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag="22">
                        </TD>
                    </TR>
                    <TR>
                        <TD CLASS="TDFAMILY_TITLE" NOWRAP>생년월일</TD>
                        <TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" ID="txtbirt" NAME="txtbirt" ALT="생년월일" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22D"  ondblclick="VBScript:Call OpenCalendar('txtbirt',3)">
                           	<INPUT TYPE="RADIO" NAME="txtso_lu_cd" tag="22" CHECKED ID="txtso_lu_cd1" VALUE=1 STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtso_lu_cd1">양력</LABEL>
    					    <INPUT TYPE="RADIO" NAME="txtso_lu_cd" tag="22" ID="txtso_lu_cd2" VALUE=2 STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtso_lu_cd2">음력</LABEL>
    					    <INPUT TYPE="hidden" name="txtso_lu_cdv"  ID="txtso_lu_cdv">
                        </TD>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>기념일</TD>
	            		<TD CLASS="TDFAMILY">
	            		    <SELECT NAME="txtMemo_cd" ALT=기념일구분 STYLE="WIDTH: 100px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
                            <INPUT CLASS="SINPUTTEST_STYLE" ID="txtMemo_dt" NAME="txtMemo_dt" ALT="기념일" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22D" ondblclick="VBScript:Call OpenCalendar('txtMemo_dt',3)">
                        </TD>
                    </TR>
                    <TR>
	            		<TD CLASS=TDFAMILY_TITLE NOWRAP>결혼여부</TD>
	            		<TD CLASS=TDFAMILY align=left>
	            		    <SELECT NAME="txtMarry_Cd" ALT=결혼여부 STYLE="WIDTH: 100px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
	            		</TD>
	            		<TD CLASS="TDFAMILY_TITLE" NOWRAP>주거구분</TD>
	            		<TD CLASS="TDFAMILY">
	            		    <SELECT NAME="txtHouse_Cd" ALT=주거구분 STYLE="WIDTH: 100px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
                        </TD>
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP>본적</TD>
	                    <TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDomi" ALT="본적" TYPE="Text" MAXLENGTH=20 SiZE=22 tag=22>
                        </TD>
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP rowspan=2>주민등록지</TD>
	                    <TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtZip_cd" ALT="주민등록지우편번호" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=22><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZip(1)">
                        </TD>      
                    </TR>
                    <TR>

	                    <TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtAddr" ALT="주민등록지" TYPE="Text" MAXLENGTH=128 SiZE=80 tag=22>
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP rowspan=2>현주소</TD>
	                    <TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCurr_zip_cd" ALT="현주소우편번호" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=22><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZip(2)">
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY2" colspan=3>
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCurr_addr" ALT="현주소" TYPE="Text" MAXLENGTH=128 SiZE=80 tag=22>
                        </TD>      
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP>전화번호</TD>
	                    <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTel_no" ALT="전화번호" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=22></TD>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP>비상연락번호</TD>
	                    <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEm_tel_no" ALT="비상연락번호" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=22></TD>
                    </TR>
                    <TR>
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP>E-Mail</TD>
	                    <TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEmail_addr" ALT="E-Mail" TYPE="Text" MAXLENGTH=30 SiZE=30 tag=22>
                        </TD>      
	                    <TD CLASS="TDFAMILY_TITLE" NOWRAP>핸드폰</TD>
	                    <TD CLASS="TDFAMILY">
                            <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtHand_tel_no" ALT="핸드폰" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=22>
                        </TD>      
                    </TR>
                    <TR>
	            	    <TD CLASS="TDFAMILY_TITLE" NOWRAP>신장</TD>
	            	    <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtHgt" ALT="신장" TYPE="Text" MAXLENGTH=5 SiZE=5 tag="22">&nbsp;Cm</TD>
	            	    <TD CLASS="TDFAMILY_TITLE" NOWRAP>체중</TD>
	            	    <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtWgt" ALT="체중" TYPE="Text" MAXLENGTH=5 SiZE=5 tag="22">&nbsp;Kg</TD>
                    </TR>
                    <TR>
	            	    <TD CLASS="TDFAMILY_TITLE" NOWRAP>혈액형</TD>
	            	    <TD CLASS="TDFAMILY"><SELECT NAME="txtBlood_type1" ALT="혈액형1" STYLE="WIDTH: 70px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>&nbsp;형
	            	                         <SELECT NAME="txtBlood_type2" ALT="혈액형2" STYLE="WIDTH: 70px" TAG="22"><OPTION VALUE=""></OPTION></SELECT></TD>
	            	    <TD CLASS="TDFAMILY_TITLE" NOWRAP>색맹</TD>
	            	    <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="22" NAME="txtDalt_type"></TD>
                    </TR>

                    <TR>
	            	    <TD CLASS="TDFAMILY_TITLE" NOWRAP>시력</TD>
	            	    <TD CLASS="TDFAMILY">&nbsp;좌<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEyesgt_left" ALT="시력(좌)" TYPE="Text" MAXLENGTH=4 SiZE=5 tag="22">
                                             &nbsp;우<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEyesgt_right" ALT="시력(우)" TYPE="Text" MAXLENGTH=4 SiZE=5 tag="22"></TD>
                        <TD CLASS="TDFAMILY_TITLE"></TD>
                        <TD CLASS="TDFAMILY"></TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
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
