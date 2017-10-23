<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 인사기본자료관리 
'*  3. Program ID           : h2009oa1
'*  4. Program Name         : 인사기록대장출력 
'*  5. Program Desc         : 인사기록대장출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->

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
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    ' 직종 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(frm1.cboOcpt_type, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables
    Call InitComboBox 
                                          
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetToolbar("1000000000000111")
    
    frm1.txtEmp_no.focus 
    
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If
    If  txtFr_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
    
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
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
' Name : OpenEmptName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
	End If

    arrParam(2) = lgUsrIntcd
    
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
		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
	arrParam(1) = ""
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
             Case Else
        End Select	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_internal_cd.value = arrRet(2) 
               .txtTo_dept_cd.focus
             Case Else
        End Select
	End With
End Function       		
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
   
	Dim emp_no, ocpt_type
	Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
	
	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
		frm1.cboOcpt_type.value = ""
	End if	

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
        Exit Function
    End If
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   

	strUrl =          "emp_no|"      & emp_no
	strUrl = strUrl & "|ocpt_cd|"    & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd
	
	ObjName = AskEBDocumentName("h2009oa1","ebr")
	call FncEBRPrint(EBAction,ObjName, strUrl)
	
End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
   
	Dim strUrl
	Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim emp_no, ocpt_type
	Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
	dim ObjName
   	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
		frm1.cboOcpt_type.value = ""
	End if	

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange()  then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If
       
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   
   
	strUrl = strUrl & "emp_no|" & emp_no
	strUrl = strUrl & "|ocpt_cd|" & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd
	
	ObjName = AskEBDocumentName("h2009oa1","ebr")
	call FncEBRPreview(ObjName, strUrl)
	
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

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'=======================================================================================================
'   Event Name : txtTo_retire_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtEmp_no_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub     

Sub txtFr_dept_cd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtto_dept_cd_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>인사기록대장출력</font></td>
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
								<TD CLASS="TD5" NOWRAP>대상자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtEmp_no" NAME="txtEmp_no" SIZE=13 MAXLENGTH=13  tag="11XXXU" ALT="대상자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName(1)">
								                       <INPUT TYPE="Text" NAME="txtName" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="대상자코드명"></TD>
							</TR>
							    <TD CLASS=TD5 NOWRAP>직종</TD>
						  		<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOcpt_type" ALT="직종" STYLE="WIDTH: 100px" TAG="11XXXU"><OPTION VALUE=""></OPTION></SELECT>		
							<TR>			
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40"  tag="14XXXU">부터
		                                             <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">
		                    </TR>
		                    <TR>    
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE="10"  ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE="20"  ALT ="Order ID" tag="14XXXU">까지
    			                                     <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7"  tag="14XXXU"></TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()" flag=1>미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" flag=1>인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST"> 
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>


