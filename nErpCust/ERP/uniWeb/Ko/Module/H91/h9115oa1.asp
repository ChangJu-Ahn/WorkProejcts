<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 연말정산관리 
'*  3. Program ID           : h9115oa1
'*  4. Program Name         : 소득세 납세필증명서 
'*  5. Program Desc         : 소득세 납세필증명서 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/06/14
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : BongKyu, Song
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lsInternal_cd
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtPresentDt.focus
    
	frm1.txtPresentDt.text = UNIConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat, parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtPresentDt, parent.gDateFormat, 1)

	frm1.txtBasYymm.text = UNIConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat, parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtBasYymm, parent.gDateFormat, 2)

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("P", "H", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
        
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")

    Call InitVariables                                        
    Call SetDefaultVal

    Call SetToolbar("1000000000000111")
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function
'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmpName()
'	Description : Item Popup에서 Return되는 값 setting(grid외에서 사용)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
' Name : FncBtnPrint
' Desc : developer describe this line 
'========================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

	dim emp_no, presentdt, basyymm, presentnm, officenm, cntsum, reason
	Dim ObjName
	StrEbrFile = "h9115oa1"

	emp_no    = frm1.txtEmp_no.value
    presentdt = frm1.txtPresentDt.Year & Right("0" & frm1.txtPresentDt.Month, 2) & Right("0" & frm1.txtPresentDt.Day, 2) 
    basyymm   = frm1.txtBasYymm.Year & Right("0" & frm1.txtBasYymm.Month, 2)
	presentnm = frm1.txtPresentNm.value 'Trim(UCase(frm1.txtPresentNm.value))
	officenm  = frm1.txtOfficeNm.value 'Trim(UCase(frm1.txtOfficeNm.value))
	cntsum    = frm1.txtCntSum.value 'Trim(UCase(frm1.txtCntSum.value))
	reason    = frm1.txtReason.value 'Trim(UCase(frm1.txtReason.value))

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

    If txtEmp_no_Onchange() Then                                                'enter key 로 조회시 시작부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

	if presentnm = "" then
		presentnm = "%"
		frm1.txtPresentNm.value = ""
	End if	

	if officenm = "" then
		officenm = "%"
		frm1.txtOfficeNm.value = ""
	End if	

	if reason = "" then
		reason = "%"
		frm1.txtReason.value = ""
	End if	
					
	strUrl = "Emp_no|" & emp_no
	strUrl = strUrl & "|PresentDt|" & presentdt
	strUrl = strUrl & "|Basyymm|" & basyymm
	strUrl = strUrl & "|PresentNm|" & presentnm
	strUrl = strUrl & "|OfficeNm|" & officenm
	strUrl = strUrl & "|CntSum|" & cntsum
	strUrl = strUrl & "|Reason|" & reason

   	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , strUrl)
		
End Function


'========================================================================================================
' Name : BtnPreview
' Desc : This function is related to Preview Button
'========================================================================================================
Function BtnPreview() 
'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim ObjName		
	dim emp_no, presentdt, basyymm, presentnm, officenm, cntsum, reason
	
	StrEbrFile = "h9115oa1"
	
	emp_no    = frm1.txtEmp_no.value
    presentdt = frm1.txtPresentDt.Year & Right("0" & frm1.txtPresentDt.Month, 2) & Right("0" & frm1.txtPresentDt.Day, 2) 
    basyymm   = frm1.txtBasYymm.Year & Right("0" & frm1.txtBasYymm.Month, 2)
	presentnm = Trim(UCase(frm1.txtPresentNm.value))
	officenm  = Trim(UCase(frm1.txtOfficeNm.value))
	cntsum    = Trim(UCase(frm1.txtCntSum.value))
	reason    = Trim(UCase(frm1.txtReason.value))

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

    If txtEmp_no_Onchange() Then                                                'enter key 로 조회시 시작부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

	if presentnm = "" then
		presentnm = "%"
		frm1.txtPresentNm.value = ""
	End if	

	if officenm = "" then
		officenm = "%"
		frm1.txtOfficeNm.value = ""
	End if	

	if reason = "" then
		reason = "%"
		frm1.txtReason.value = ""
	End if	
					
	strUrl = "Emp_no|" & emp_no
	strUrl = strUrl & "|PresentDt|" & presentdt
	strUrl = strUrl & "|Basyymm|" & basyymm
	strUrl = strUrl & "|PresentNm|" & presentnm
	strUrl = strUrl & "|OfficeNm|" & officenm
	strUrl = strUrl & "|CntSum|" & cntsum
	strUrl = strUrl & "|Reason|" & reason
	
   	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)
		
End Function

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                              strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
' Name : txtPresentdt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPresentdt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPresentdt.Action = 7
		frm1.txtPresentdt.focus
	End If
End Sub

'========================================================================================================
' Name : txtBasYymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtBasYymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBasYymm.Action = 7
		frm1.txtBasYymm.focus
	End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>소득세납세필증명서</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>   
							<TR>
								<TD CLASS=TD5 NOWRAP>발급일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9115oa1_txtPresentdt_txtPresentDt.js'></script>&nbsp;</TD>
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>기준년월</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/h9115oa1_fpDateTime1_txtBasYymm.js'></script>
								</TD>															
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>제출처</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtPresentNm" ALT="제출처" TYPE="Text" SiZE=20 MAXLENGTH=14 tag="21XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세무서명</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtOfficeNm" ALT="세무서명" TYPE="Text" SiZE=20 MAXLENGTH=20 tag="21XXXU"></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>신청인</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="신청인" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">&nbsp;
								                     <INPUT NAME="txtName" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="14XXXU"></TD>	
						    </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>소요수량</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtCntSum" ALT="소요수량" VALUE="1" TYPE="Text" SiZE=5 MAXLENGTH=2 tag="11XXXU">통</td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사용목적</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtReason" ALT="사용목적" TYPE="Text" SiZE=40 MAXLENGTH=40 tag="21XXXU"></td>
							</TR>
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>

</BODY>
</HTML>

