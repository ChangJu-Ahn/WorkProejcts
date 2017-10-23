<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 연월차계산 
*  3. Program ID           : H9201ba1
*  4. Program Name         : H9201ba1
*  5. Program Desc         : 연말정산관리/연월차계산 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/04/25
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     : songbongkyu
* 10. Modifier (Last)      : Lee Si Na
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID = "H9201bb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
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

	frm1.txtyear_yymm.focus
	Call ExtractDateFrom("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	Call ggoOper.FormatDate(frm1.txtyear_yymm, parent.gDateFormat, 2)
	frm1.txtyear_yymm.Year	= strYear
	frm1.txtyear_yymm.Month = strMonth
	frm1.txtyear_yymm.Day	= strDay
	
	Call ggoOper.FormatDate(frm1.txtRetire_stdt, parent.gDateFormat, 1)	
	frm1.txtRetire_stdt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtRetire_enddt.Text = frm1.txtRetire_stdt.Text
	
	Call ggoOper.SetReqAttr(frm1.txtRetire_stdt, "Q")
	Call ggoOper.SetReqAttr(frm1.txtRetire_enddt, "Q")		
	Call ggoOper.SetReqAttr(frm1.txtdilig_cd, "Q")
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A( "B", "H","NOCOOKIE","BA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = 'H0046' AND MINOR_CD NOT IN('$','3') ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtyear_type, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("DILIG_CD, DILIG_NM ","HCA010T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtdilig_cd, lgF0, lgF1, Chr(11))
    
End Sub
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
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value				' Name Cindition
	End If
    arrParam(2) = ""
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
			Set gActiveElement = document.ActiveElement

			lgBlnFlgChgValue = False
		End With
	End If	
			
End Function
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
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Dim strVal
	Dim strDate
    Dim strYear,strMonth,strDay
	Dim IntRetCD
	
	On Error Resume Next             

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if  txtEmp_no_OnChange() then
		Exit Function
	end if
	
	strDate		= UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtyear_yymm.Year,Right("0" & frm1.txtyear_yymm.Month,2),"01")	
    IF  FuncAuthority("@", UniConvDateToYYYYMMDD(strDate,parent.gDateFormat,""), parent.gUsrID) = "N" THEN
        '연월차 마감처리된 지급월 입니다.
        Call DisplayMsgbox("800304","X","X","X")
        exit function
    END IF

    IF  frm1.txtyear_type.value = "2" THEN  '월차이면 
        If  CommonQueryRs(" prov_type "," HDA140T ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
     
            if  Replace(lgF0, Chr(11), "") <> "3" then
                '월차수당기준등록에서 지급방법을 년말정산으로 등록하세요.", Exclamation!)
                Call DisplayMsgbox("800413","X","X","X")
                exit function
            end if
        else
            call msgbox("hda140t table error")
            exit function
        end if
    Else
        If  CommonQueryRs(" * "," HDA100T "," allow_cd='Y01' and  flag ='2' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
			Call DisplayMsgbox("971012","X","연차수당기준등록에서 연차발생근태코드가 등록되었는지","X")
			exit function
        end if    
    end if
    

	IntRetCD = DisplayMsgbox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	if LayerShowHide(1) = false then
	    Exit Function
	end if

	ExeReflect = False                                                          '⊙: Processing is NG
    
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0006
	strVal = strVal & "&txtyear_yymm=" & frm1.year_yymm.year & right("0" & frm1.year_yymm.month,2)

	strVal = strVal & "&txtyear_type=" & frm1.txtyear_type.value
	
	if  frm1.tax_calc1.checked  then
	    strVal = strVal & "&txttax_calc=Y"
	else
	    strVal = strVal & "&txttax_calc=N"
	end if

	strVal = strVal & "&txtdilig_cd=" & frm1.txtdilig_cd.value
	
	strVal = strVal & "&txtPay_cd=" & frm1.txtPay_cd.value
    strVal = strVal & "&txtEmp_no=" & frm1.txtEmp_no.value

    IF  frm1.txtyear_type.value = "2" THEN  '월차이면 
        If  CommonQueryRs(" MAX(allow_cd) "," HDA140T ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
            strVal = strVal & "&txtallow_cd=" & Replace(lgF0, Chr(11), "")
        else
            strVal = strVal & "&txtallow_cd=" & ""
        end if
    ElseIF  frm1.txtyear_type.value = "1" THEN  '연차이면 
        If  CommonQueryRs(" MAX(allow_cd) "," HDA150T ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
            strVal = strVal & "&txtallow_cd=" & Replace(lgF0, Chr(11), "")
        else
            strVal = strVal & "&txtallow_cd=" & ""
        end if
    end if

	If frm1.chk_retire_flag.checked = True Then	
		strVal = strVal & "&txtRetireFlag=Y"
		Call ExtractDateFrom(frm1.txtRetire_stdt.Text, parent.gDateFormat, parent.gComDateType, strYear, strMonth, strDay)
		strVal = strVal & "&txtRetire_stdt=" & strYear & right("0" & strMonth, 2) & right("0" & strDay, 2)
		Call ExtractDateFrom(frm1.txtRetire_enddt.Text, parent.gDateFormat, parent.gComDateType, strYear, strMonth, strDay)
		strVal = strVal & "&txtRetire_enddt=" & strYear & right("0" & strMonth, 2) & right("0" & strDay, 2)
	Else
		strVal = strVal & "&txtRetireFlag=N"
		strVal = strVal & "&txtRetire_stdt=19000101" 
		strVal = strVal & "&txtRetire_enddt=19000101"
	End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG

End Function


Function FuncAuthority(Pay_gubun, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"    
    IntRetCD = CommonQueryRs("close_type, close_dt, emp_no","hda270t","org_cd=" & FilterVar("1", "''", "S") & "  and pay_gubun=" & FilterVar("Z", "''", "S") & "  and pay_type= " & FilterVar(Pay_gubun, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  IntRetCD = false then
        strRet = "Y"
    else
        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF  UniConvDateToYYYYMMDD(Replace(lgF1,Chr(11),""),parent.gServerDateFormat,"") <= Pay_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    IF  UniConvDateToYYYYMMDD(Replace(lgF1,Chr(11),""),parent.gServerDateFormat,"") < Pay_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT
        
    end if

    FuncAuthority = strRet

End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	call DisplayMsgbox("990000","X","X","X")
	window.status = "작업 완료"
	frm1.txtyear_yymm.focus
End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgbox("800161","X","X","X")
	window.status = "작업 완료"
	frm1.txtyear_yymm.focus
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
		With frm1
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
			lgBlnFlgChgValue = False
		End With

	End If	
			
End Function

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtyear_yymm, parent.gDateFormat, 2)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
	Call InitComboBox()

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
Sub txtyear_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtyear_yymm.Action = 7
		frm1.txtyear_yymm.focus
	End If
End Sub
Sub txtRetire_stdt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetire_stdt.Action = 7
		frm1.txtRetire_stdt.focus
	End If
End Sub
Sub txtRetire_enddt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtRetire_enddt.Action = 7
		frm1.txtRetire_enddt.focus
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

Function chk_retire_flag_onclick()

	If frm1.chk_retire_flag.checked = True Then	
		 Call ggoOper.SetReqAttr(frm1.txtRetire_stdt, "N")
		 Call ggoOper.SetReqAttr(frm1.txtRetire_enddt, "N")		
	Else
		 Call ggoOper.SetReqAttr(frm1.txtRetire_stdt, "Q")
		 Call ggoOper.SetReqAttr(frm1.txtRetire_enddt, "Q")				
	End If
End Function

Function txtYear_type_onChange()

	If frm1.txtYear_type.value = "1" Then
		Call ggoOper.SetReqAttr(frm1.txtdilig_cd, "Q")	
	Else
		Call ggoOper.SetReqAttr(frm1.txtdilig_cd, "N")	
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연월차계산</font></td>
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
								<TD CLASS=TD5 NOWRAP>연월차년월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=year_yymm NAME="txtyear_yymm" CLASS=FPDTYYYYMM tag="12X1" ALT="연월차년월" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>퇴직구분</TD>
								<TD CLASS=TD6 NOWRAP>퇴직자	<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chk_retire_flag" ID="chk_retire_flag"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계산퇴직기간</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtRetire_stdt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="퇴직계산시작일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
											  &nbsp; ~ &nbsp;
											  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtRetire_enddt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="퇴직계산종료일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
    						<TR>
								<TD CLASS=TD5 NOWRAP>연월차구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtYear_type" ALT="연월차구분" STYLE="WIDTH: 133px" tag="12"></SELECT></TD>
							</TR>
    						<TR>
								<TD CLASS=TD5 NOWRAP>관련근태</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtdilig_cd" ALT="관련근태" STYLE="WIDTH: 133px" tag="12"></SELECT></TD>
							</TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>세액계산여부</TD>
				                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txttax_calc" TAG="2X" VALUE="Y" CHECKED ID="tax_calc1"><LABEL FOR="tax_calc1">Y</LABEL>&nbsp;
				                                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txttax_calc" TAG="2X" VALUE="N" ID="tax_calc2"><LABEL FOR="tax_calc2">N</LABEL></TD>
						    </TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
	                    		<TD CLASS="TD6" NOWRAP>
                					<SELECT Name="txtPay_cd" ALT="급여구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT>
		                    	</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH="13" SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH="30" SIZE=20 tag=14XXXU></TD>	
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
					<TD Width=10>&nbsp</TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
