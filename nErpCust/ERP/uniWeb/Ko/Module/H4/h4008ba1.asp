<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 월근태집계 
*  3. Program ID           : H4008ba1
*  4. Program Name         : H4008ba1
*  5. Program Desc         : 근태관리/월근태집계 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/01
*  8. Modified date(Last)  : 2003/06/11
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H4008bb1.asp"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
    Call ggoOper.FormatDate(frm1.txtWk_yymm, parent.gDateFormat, 2)
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtWk_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtWk_yymm.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","BA") %>
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtWarrentNo.Value & parent.gColSep       'You Must append one character(parent.gColSep)
   Else
      lgKeyStream = Frm1.txtMajorCd.Value & parent.gColSep         'You Must append one character(parent.gColSep)
   End If   
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet

	Call InitComboBox()

	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 

	frm1.txtBiz_area.focus
			
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
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Function

'========================================================================================================
'	Name : OpenBiz_area()
'	Description : 
'========================================================================================================
Function OpenBiz_area()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "b_biz_area"					' TABLE 명칭 
	arrParam(2) = frm1.txtBiz_area.value		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 
	
    arrField(0) = "biz_area_cd"					' Field명(0)
    arrField(1) = "biz_area_nm"					' Field명(1)
    
    arrHeader(0) = "사업장코드"				' Header명(0)
    arrHeader(1) = "사업장명"				' Header명(1)

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBiz_area.focus	
		Exit Function
	Else
		Call SetBiz_area(arrRet)
	End If	
	
End Function

Function SetBiz_area(arrRet)
	Dim strFlag
	Dim lRow
	
	With frm1
        .txtBiz_area.value = arrRet(0)
		.txtBiz_area_Nm.value = arrRet(1)
		.txtBiz_area.focus
	End With
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
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	End If
    arrParam(2) = ""
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


'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 

	Dim strVal
	Dim strWkYear
	Dim strWkMonth
	Dim strWkYYYYMM
	Dim strYYYYMMDD
	Dim IntRetCD
	Dim strYear,strMonth,strDay

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If
	if txtEmp_no_Onchange() then
		Exit Function
	end if
	if txtBiz_area_Onchange() then
		Exit Function
	end if		

    strWkYYYYMM = frm1.txtWk_yymm.Year & Right("0" & frm1.txtWk_yymm.Month, 2)
	strYYYYMMDD = UniConvDateAToB(frm1.txtWk_yymm,parent.gDateFormatYYYYMM,parent.gServerDateFormat)

    IF  FuncAuthority("#", strYYYYMMDD, parent.gUsrID) = "N" THEN
        Call DisplayMsgBox("800291","X","X","X")         '"근태 마감처리된 일 입니다."
        Call BtnDisabled(0)
        Exit function
    END IF

    If  frm1.txtPay_cd.value = "" Then
    
        IntRetCD = CommonQueryRs("dilig_dd ","hda000t","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)          'SELECT dilig_dd FROM hda000t;
        If  IntRetCD = false Then
		    Call DisplayMsgBox("800412","X","X","X")              '"회사RULE등록에서 근태기준일을 입력하십시오."
            Call BtnDisabled(0)
            Exit function
        End if
    Else

        IntRetCD = CommonQueryRs("COUNT(*) ","hda190t","pay_cd =  " & FilterVar(frm1.txtPay_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)          'SELECT COUNT(*) FROM hda190t WHERE pay_cd = :ls_pay_cd ;
        If  CInt(Replace(lgF0,Chr(11),"")) = 0 Then
		    Call DisplayMsgBox("800411","X","X","X")              '"기본급산출기준식에 해당 급여구분의 근태기준을 입력하십시오."
            Call BtnDisabled(0)
            Exit function
        End if
        
    End if

    ' 근무칼렌다 생성여부 체크 
    IntRetCD = CommonQueryRs("COUNT(*) ","hca020t","org_cd =  " & FilterVar(frm1.txtBiz_area.Value , "''", "S") & " and date =  " & FilterVar(strYYYYMMDD , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  CInt(Replace(lgF0,Chr(11),"")) = 0 Then
        Call DisplayMsgBox("800453","X","X","X")             
        Call BtnDisabled(0)
        Exit function
    End if        

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If LayerShowHide(1) = False then
    	Call BtnDisabled(0)
    	Exit Function 
    End if

	strVal = BIZ_PGM_ID & "?txtMode="     & parent.UID_M0006
	strVal = strVal     & "&txtbiz_area=" & frm1.txtbiz_area.value
	strVal = strVal     & "&txtPay_yymm=" & strWkYYYYMM
	strVal = strVal     & "&txtPay_cd="   & frm1.txtPay_cd.value

    ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	call DisplayMsgBox("990000","X","X","X")
	frm1.txtBiz_area.focus	
End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgBox("800161","X","X","X")
	frm1.txtBiz_area.focus	
End Function

Function FuncAuthority(Pay_gubun, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"
    IntRetCD = CommonQueryRs("close_type, CONVERT(CHAR(10),close_dt, 20), emp_no","hda270t","org_cd=" & FilterVar("1", "''", "S") & "  and pay_gubun=" & FilterVar("Z", "''", "S") & "  and pay_type= " & FilterVar(Pay_gubun, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  IntRetCD = false then
        strRet = "Y"
    else

        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF  Replace(lgF1, Chr(11), "") <= Pay_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    IF  Replace(lgF1, Chr(11), "") < Pay_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT
        
    end if

    FuncAuthority = strRet

End Function
'======================================================================================================
'   Event Name : txtWk_yymm_DblClick
'   Event Desc :
'=======================================================================================================
Sub txtWk_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtWk_yymm.Action = 7
		frm1.txtWk_yymm.focus
	End If
End Sub

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

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
    
End Function

'========================================================================================================
'   Event Name : txtBiz_area_Onchange             
'   Event Desc :
'========================================================================================================
Function txtBiz_area_Onchange()
    Dim IntRetCd

    If  frm1.txtBiz_area.value = "" Then
		frm1.txtBiz_area_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" Biz_area_nm "," b_Biz_area "," Biz_area_cd =  " & FilterVar(frm1.txtBiz_area.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        If  IntRetCd = false then
			Call DisplayMsgBox("970000","X","사업장코드","X")
			frm1.txtBiz_area_nm.value = ""
            frm1.txtBiz_area.focus
            Set gActiveElement = document.ActiveElement       
			txtBiz_area_Onchange  = true
			exit Function
        Else
			frm1.txtBiz_area_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if
    
End Function 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월근태집계</font></td>
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
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBiz_area" ALT="사업장" TYPE="Text" MAXLENGTH=10 SIZE=10 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz_area">&nbsp;<INPUT NAME="txtBiz_area_nm" TYPE="Text" MAXLENGTH=30 SIZE=20  tag="14XXXU"></TD>	
							    </TR>
						    	<TR>
									<TD CLASS=TD5 NOWRAP>작업년월</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h4008ba1_txtWk_yymm_txtWk_yymm.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>급여구분</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="txtPay_cd" ALT="급여구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT>
		                        	</TD>
								</TR>			
								<TR>
									<TD CLASS=TD5 NOWRAP>대상자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH=13 SIZE=13 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20  tag="14XXXU"></TD>	
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


