<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 퇴직금계산 
*  3. Program ID           : Ha104ba1
*  4. Program Name         : Ha104ba1
*  5. Program Desc         : 퇴직정산관리/퇴직금계산 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'=======================================================================================================
Const BIZ_PGM_ID = "Ha104bb1.asp"
Dim StartDate
Dim IsOpenPop          

StartDate	= "<%=GetSvrDate%>"

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
Dim strYear,strMonth,strDay
	frm1.txtretire_yyyy.focus
	Call ggoOper.FormatDate(frm1.txtretire_yyyy, parent.gDateFormat, 3)    
	Call ExtractDateFrom(StartDate,parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)	
	frm1.txtretire_yyyy.Year	= strYear
	frm1.txtretire_yyyy.Month	= strMonth
	frm1.txtretire_yyyy.Day		= strDay
	
	frm1.txtretire_strt_dt.text = UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtretire_end_dt.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "BA") %>
End Sub

'======================================================================================================
'   Event Name : txtEmp_no_OnChange
'   Event Desc : 사번(성명)이 변경될 경우 
'=======================================================================================================
Function txtEmp_no_OnChange()

    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

	frm1.txtName.value = ""
           
    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_OnChange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if

End Function

'======================================================================================================
' Function Name : FuncAuthority
' Function Desc : 시스템 마감 체크 
'=======================================================================================================
Function FuncAuthority(Pay_gubun, str_retire_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"
    IntRetCD = CommonQueryRs("close_type, CONVERT(CHAR(10),close_dt, 20), emp_no","hda270t","org_cd='1' and pay_gubun='Z' and pay_type=" & FilterVar(Pay_gubun, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       
    
    if  IntRetCD = false then
        strRet = "Y"
    else

        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF  Replace(lgF1, Chr(11), "") <= str_retire_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    IF  Replace(lgF1, Chr(11), "") < str_retire_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT

    end if

    FuncAuthority = strRet

End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Dim strVal
	Dim IntRetCD
	Dim strYear,strMonth,strDay
    Dim str_emp_no , str_strt_dt , str_end_dt , intCnt ,str_retire_yymmdd

	On Error Resume Next
	
	ExeReflect = False

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if  txtEmp_no_OnChange() then
		Exit Function
	end if

    str_emp_no = frm1.txtEmp_no.value
    str_strt_dt = frm1.txtretire_strt_dt.text
    str_end_dt  = frm1.txtretire_end_dt.text

    IF frm1.txtEmp_no.value = "" Then
		str_emp_no = "%"
    End IF

	intCnt = 0
	
	If 	CommonQueryRs(" count(*) , CONVERT(CHAR(10), min(RETIRE_DT) , 20) "," hga040t ","emp_no LIKE " &  FilterVar(str_emp_no, "''", "S")  & " AND retire_dt BETWEEN " & FilterVar(str_strt_dt, "''", "S") & " AND " & FilterVar(str_end_dt, "''", "S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
	    intCnt = CInt(Trim(Replace(lgF0, Chr(11), "")))

		IF intCnt > 0 Then
		    str_retire_yymmdd = Trim(Replace(lgF1, Chr(11), ""))

			IF  FuncAuthority("$", str_retire_yymmdd, Parent.gUsrID) = "N" THEN
			    Call DisplayMsgBox("800807","X","X","X")
			    Call BtnDisabled(0)
			    exit function
			END IF
        END If

	END If
		    
	IF CompareDateByFormat(frm1.txtretire_strt_dt.text, frm1.txtretire_end_dt.text, frm1.txtretire_strt_dt.Alt, frm1.txtretire_end_dt.Alt,"800279",	frm1.txtretire_strt_dt.UserDefinedFormat, parent.gComDateType, True) = False Then
        frm1.txtretire_strt_dt.focus()
        exit function
    END IF

	IntRetCD = DisplayMsgbox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	End If
	
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtretire_yyyy=" & frm1.txtretire_yyyy.Text		
	strVal = strVal & "&txtretire_strt_dt=" & frm1.txtretire_strt_dt.Text	
	strVal = strVal & "&txtretire_end_dt=" & frm1.txtretire_end_dt.Text

	if  frm1.txtcalcu_logic1.checked = true then
	    strVal = strVal & "&txtcalcu_logic=1"
	elseif  frm1.txtcalcu_logic2.checked = true then
	    strVal = strVal & "&txtcalcu_logic=2"
	elseif  frm1.txtcalcu_logic3.checked = true then
	    strVal = strVal & "&txtcalcu_logic=3"
	elseif  frm1.txtcalcu_logic4.checked = true then
	    strVal = strVal & "&txtcalcu_logic=4"
	end if

	if  frm1.txtpay_logic1.checked = true then
	    strVal = strVal & "&txtpay_logic=M"
	else
	    strVal = strVal & "&txtpay_logic=D"
	end if


    ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value



	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)

End Function


'======================================================================================================
' Function Name : ExeDelete
' Function Desc : 
'=======================================================================================================
Function ExeDelete()
	Dim strVal
	Dim IntRetCD
	Dim strYear,strMonth,strDay
    Dim str_emp_no , str_strt_dt , str_end_dt , intCnt , intCnt2 ,str_retire_yymmdd , str_retire_yyyy
    Dim strType , strpay_cd

	On Error Resume Next
	
	ExeReflect = False

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if  txtEmp_no_OnChange() then
		Exit Function
	end if

    str_emp_no = frm1.txtEmp_no.value
    str_strt_dt = frm1.txtretire_strt_dt.text
    str_end_dt  = frm1.txtretire_end_dt.text
    str_retire_yyyy = frm1.txtretire_yyyy.text
 
		   
    IF frm1.txtEmp_no.value = "" Then
		str_emp_no = "%"

		strpay_cd  = "*"
	    Call  CommonQueryRs(" MINOR_NM "," H_PAY_CD "," MINOR_CD = " & FilterVar(strpay_cd, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    strpay_cd = Trim(Replace(lgF0, Chr(11), ""))
    Else
 		Call CommonQueryRs("pay_cd","hdf020t","emp_no =" & FilterVar(str_emp_no, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		strpay_cd = Trim(Replace(lgF0, Chr(11), ""))
		
 		Call CommonQueryRs("MINOR_NM ","B_MINOR","MAJOR_CD = 'H0005'AND MINOR_CD = " & FilterVar(strpay_cd, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		strpay_cd = Trim(Replace(lgF0, Chr(11), ""))    
    End IF


	intCnt = 0
	If 	CommonQueryRs(" count(*) , CONVERT(CHAR(10), min(RETIRE_DT) , 20) "," hga040t ","emp_no LIKE " & FilterVar(str_emp_no, "''", "S")  & " AND retire_dt BETWEEN " & FilterVar(str_strt_dt, "''", "S") & " AND " &  FilterVar(str_end_dt, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
	    intCnt = CInt(Trim(Replace(lgF0, Chr(11), "")))

		IF intCnt > 0 Then
		    str_retire_yymmdd = Trim(Replace(lgF1, Chr(11), ""))

			IF  FuncAuthority("$", str_retire_yymmdd, Parent.gUsrID) = "N" THEN
			    Call DisplayMsgBox("800807","X","X","X")
			    Call BtnDisabled(0)
			    exit function
			END IF
        END If

	END If
	
	intCnt2 = 0
	If CommonQueryRs(" COUNT(*) "," hga070t ","emp_no LIKE " & FilterVar(str_emp_no, "''", "S")  & " AND retire_dt BETWEEN " &  FilterVar(str_strt_dt, "''", "S") & " AND " & FilterVar(str_end_dt, "''", "S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
	    intCnt2 = CInt(Trim(Replace(lgF0, Chr(11), "")))
	End if

	If intCnt2 = 0 Then
		Call DisplayMsgbox("800161","X","X","X")
		Exit Function
	End If

    Call CommonQueryRs(" allow_nm "," hda010t "," allow_cd = 'R01' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    strType = Trim(Replace(lgF0, Chr(11), ""))

'	IntRetCD = DisplayMsgBox("800805", Parent.VB_YES_NO , "(" & str_strt_dt & " ~ " & str_end_dt & ") "& strpay_cd , strType & " " & intCnt2)  '☜: Data is changed.  Do you want to display it? 
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	If IntRetCD <> vbYes Then
	    lgBlnFlgChgValue = False
	    Call BtnDisabled(0)
		Exit Function
	End If

	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	
	Call BtnDisabled(1)
	

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
	strVal = strVal & "&txtretire_strt_dt=" & frm1.txtretire_strt_dt.Text	
	strVal = strVal & "&txtretire_end_dt=" & frm1.txtretire_end_dt.Text

   ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 



	ExeReflect = True     
	Call BtnDisabled(0)		

End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	call DisplayMsgbox("990000","X","X","X")
	window.status = "작업 완료"
	frm1.txtRetire_yyyy.focus
End Function

Function ExeReflectNo()				            '☆: 처리할 데이타가 없습니다.
    Call DisplayMsgbox("800161","X","X","X")
	window.status = "작업 완료"
	frm1.txtRetire_yyyy.focus
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call InitVariables                                                     '⊙: Setup the Spread sheet
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : txtyear_yymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtRetire_yyyy_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetire_yyyy.Action = 7
		frm1.txtRetire_yyyy.focus
	End If
End Sub

Sub txtRetire_strt_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetire_strt_dt.Action = 7
		frm1.txtRetire_strt_dt.focus
	End If
End Sub

Sub txtRetire_end_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetire_end_dt.Action = 7
		frm1.txtRetire_end_dt.focus
	End If
End Sub

Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE,False)
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'=======================================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>퇴직금계산</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_40%>WIDTH=100%>   
                            <TR>
								<TD CLASS=TD5>계산년도</TD>
								<TD CLASS=TD6><script language =javascript src='./js/ha104ba1_fpDateTime2_txtRetire_yyyy.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>퇴직기간</TD>
								<TD CLASS=TD6><script language =javascript src='./js/ha104ba1_fpDateTime2_txtRetire_strt_dt.js'></script>&nbsp;~&nbsp;
								              <script language =javascript src='./js/ha104ba1_fpDateTime2_txtRetire_end_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH=13 SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20 tag=14XXXU></TD>
						    </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP>계산공식</TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="월평균임금*근속개월/12" CHECKED ID="txtCalcu_logic1"><LABEL FOR="txtCalcu_logic1">월평균임금*근속개월/12</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="월평균임금*근속일수/365"  ID="txtCalcu_logic2"><LABEL FOR="txtCalcu_logic2">월평균임금*근속일수/365</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="년/월/일퇴직금"  ID="txtCalcu_logic3"><LABEL FOR="txtCalcu_logic3">년/월/일퇴직금</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtCalcu_logic" TAG="1X" VALUE="일평균임금*30*근속일수/365"  ID="txtCalcu_logic4"><LABEL FOR="txtCalcu_logic4">일평균임금*30*근속일수/365</LABEL>
                            </TR>
                            <TR>
		        				<TD CLASS="TD5" NOWRAP>평균급여산정방법</TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPay_logic" TAG="1X" VALUE="월단위평균임금산정" CHECKED ID="txtPay_logic1"><LABEL FOR="txtPay_logic1">월단위평균임금산정</LABEL>
                            </TR>
        					<TR>
		        				<TD CLASS="TD5" NOWRAP></TD>
				        		<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtPay_logic" TAG="1X" VALUE="일단위평균임금산정"  ID="txtPay_logic2"><LABEL FOR="txtPay_logic2">일단위평균임금산정</LABEL>
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
					<TD Width = 10> &nbsp </TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>&nbsp;
					    <BUTTON NAME="btnExe2" CLASS="CLSSBTN" onclick="ExeDelete()"  Flag=1>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


