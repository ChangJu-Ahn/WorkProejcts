<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 연월차정산결과반영 
*  3. Program ID           : H9204ba1
*  4. Program Name         : H9204ba1
*  5. Program Desc         : 연말정산관리/연월차정산결과반영 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     : YBI
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
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H9202bb1.asp"
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
   
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    Dim strYear,strMonth,strDay

    frm1.txtycal_yy_dt.focus
    Call ggoOper.FormatDate(frm1.txtycal_yy_dt, parent.gDateFormat, 2)
    Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType,strYear,strMonth,strDay)	
    frm1.txtycal_yy_dt.Year		= strYear
    frm1.txtycal_yy_dt.Month	= strMonth
    frm1.txtycal_yy_dt.Day		= strDay
    
    Call ggoOper.FormatDate(frm1.txtycal_yymm_dt, parent.gDateFormat, 2)    
    frm1.txtycal_yymm_dt.Year	= strYear
    frm1.txtycal_yymm_dt.Month	= strMonth
    frm1.txtycal_yymm_dt.Day	= strDay
    
  	Call ggoOper.FormatDate(frm1.txtyear_yymm_dt, parent.gDateFormat, 1)	
    frm1.txtyear_yymm_dt.Year	= strYear
    frm1.txtyear_yymm_dt.Month	= strMonth
    frm1.txtyear_yymm_dt.Day	= strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("B", "H","NOCOOKIE","BA") %>
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

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD NOT IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " ," & FilterVar("P", "''", "S") & " ," & FilterVar("Q", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProv_type, lgF0, lgF1, Chr(11))
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
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
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
    arrParam(2) = ""
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
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
	Dim strVal
	Dim IntRetCD
    Dim intCnt
    Dim strYear, strMonth, strDay
    Dim strProv_yymm, strProv_type
    Dim strYcal_yymm, strYcal_Reflect_yymm
	Dim strWhere,strEmpno,strPaycd
	
	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	if  txtEmp_no_OnChange() then
		Exit Function
	end if

    If 	CommonQueryRs(" COUNT(*) "," hda010t "," allow_cd=" & FilterVar("P13", "''", "S") & " AND code_type=" & FilterVar("1", "''", "S") & "  AND pay_cd=" & FilterVar("*", "''", "S") & "  AND calcu_type=" & FilterVar("N", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
        intCnt = CInt(Replace(lgF0, Chr(11), ""))
    end if
    
    if  intCnt <= 0 then
        Call DisplayMsgbox("800427","X","X","X")
        Exit function
    end if

	If CompareDateByFormat(frm1.txtycal_yy_dt.Text,frm1.txtycal_yymm_dt.Text, frm1.txtycal_yy_dt.Alt, frm1.txtycal_yymm_dt.Alt, "970025", parent.gDateFormatYYYYMM, parent.gComDateType, True) = False Then
         Exit Function
    End IF 

    Call ExtractDateFrom(frm1.txtycal_yy_dt.Text, parent.gDateFormatYYYYMM, parent.gComDateType, strYear, strMonth, strDay)
    strYcal_yymm = strYear & right("0" & strMonth,2)
    
    Call ExtractDateFrom(frm1.txtycal_yymm_dt.Text, parent.gDateFormatYYYYMM, parent.gComDateType, strYear, strMonth, strDay)
    strYcal_Reflect_yymm = strYear & right("0" & strMonth,2)

'시스템마감정보확인하여 메세지처리 2007.04.13  800294 급여 마감처리된 지급월 입니다. 
      If CommonQueryRs(" COUNT(*) "," hda270t ", " pay_type=" & FilterVar("!", "''", "S") & " AND close_type=" & FilterVar("2", "''", "S") &  "  AND convert(varchar(6),close_dt,112) >=" & FilterVar(strYcal_Reflect_yymm, "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
        intCnt = CInt(Replace(lgF0, Chr(11), ""))
    end if

    if  intCnt > 0 then
        Call DisplayMsgbox("800294","X","X","X") 
        Exit function
    end if

    ' 연월차 결과 반영 월과 지급구분과 현재 작업하는 것과 다를 경우 메시지 처리 2003.04.09 by sbk  
    if Trim(frm1.txtemp_no.value)="" then
		strEmpno = "%"
    else
		strEmpno = Trim(frm1.txtemp_no.value)
    end if
    
    if Trim(frm1.txtPay_cd.value)="" then
		strPaycd = "%"
    else
		strPaycd = Trim(frm1.txtPay_cd.value)
    end if
    strWhere = " hfb020t.emp_no=hdf020t.emp_no and hfb020t.year_yymm=" & FilterVar(strYcal_yymm, "''", "S") 
    strWhere = strWhere & " and hfb020t.emp_no LIKE " & FilterVar(strEmpno, "''", "S")  & " and hdf020t.pay_cd LIKE " &  FilterVar(strPaycd, "''", "S")

    If 	CommonQueryRs(" distinct hfb020t.prov_yymm,hfb020t.prov_type ","  hfb020t ,hdf020t ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then

        strProv_yymm = Replace(lgF0, Chr(11), "")
        strProv_type = Replace(lgF1, Chr(11), "")

        If IsNull(strProv_yymm) OR strProv_yymm = "" Then
	        IntRetCD = DisplayMsgbox("900018",parent.VB_YES_NO,"X","X")	'작업을 수행 하시겠습니까?
	        If IntRetCD = vbNo Then
	        	Exit Function
	        End If
        Else
            If strYcal_Reflect_yymm <> strProv_yymm OR frm1.txtProv_type.value <> strProv_type Then
                If 	CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0040", "''", "S") & " AND MINOR_CD= " & FilterVar(strProv_type, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
	                IntRetCD = DisplayMsgbox("800602",parent.VB_YES_NO,strProv_yymm,Replace(lgF0, Chr(11), ""))
								'%1월 %2(으)로 이미 생성되어 있습니다. 다시 생성하시겠습니까 
	                If IntRetCD = vbNo Then
	                	Exit Function
	                End If
	            End If
	        Else         
	            IntRetCD = DisplayMsgbox("800397",parent.VB_YES_NO,"X","X") '이미 생성된 자료가 있습니다. 재생성하시겠습니까.
	
	            If IntRetCD = vbNo Then
	            	Exit Function
	            End If
            End if
        End if
    Else
        Call DisplayMsgbox("800161","X","X","X") '처리할 데이타가 없습니다.
        Exit function    
    End if

	If LayerShowHide(1) = false then
	    Exit Function
	End if
	
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0006
	strVal = strVal & "&txtycal_yy_dt=" & strYcal_yymm
	strVal = strVal & "&txtycal_yymm_dt=" & strYcal_Reflect_yymm
	strVal = strVal & "&txtyear_yymm_dt=" & frm1.year_yymm_dt.year & right("0" & frm1.year_yymm_dt.month,2) & right("0" & frm1.year_yymm_dt.day,2)
	strVal = strVal & "&txtprov_type=" & frm1.txtProv_type.value
	if  frm1.tax_calc1.checked  then
	    strVal = strVal & "&txttax_calc=Y"
	elseif frm1.tax_calc2.checked Then
	    strVal = strVal & "&txttax_calc=N" 
	else 
	    strVal = strVal & "&txttax_calc=H" 
	end if
    strVal = strVal & "&txtpay_cd=" & frm1.txtpay_cd.value

    ' Business Logic에서 emp_no check('%')
    strVal = strVal & "&txtEmp_no=" & frm1.txtemp_no.value

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG

End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	call DisplayMsgbox("990000","X","X","X")
	frm1.txtycal_yy_dt.focus
End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgbox("800161","X","X","X")
	frm1.txtycal_yy_dt.focus
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
'   Event Name : txtyear_yymm_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtycal_yy_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtycal_yy_dt.Action = 7
		frm1.txtycal_yy_dt.focus
	End If
End Sub

Sub txtYcal_yymm_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtYcal_yymm_dt.Action = 7
		frm1.txtYcal_yymm_dt.focus
	End If
End Sub

Sub txtyear_yymm_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtyear_yymm_dt.Action = 7
		frm1.txtyear_yymm_dt.focus
	End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연월차정산결과반영</font></td>
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
								<TD CLASS=TD5 NOWRAP>연월차정산년월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtYcal_yy_dt" CLASS=FPDTYYYYMM tag="12X1" ALT="연월차정산년월" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>연월차정산결과반영월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=present_dt NAME="txtYcal_yymm_dt" CLASS=FPDTYYYYMM tag="12X1" ALT="연월차정산결과반영월" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>	
							<TR>	
								<TD CLASS=TD5 NOWRAP>연월차지급일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=year_yymm_dt NAME="txtyear_yymm_dt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="연월차지급일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>연월차정산결과지급구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtProv_type" ALT="연월차구분" STYLE="WIDTH: 133px" tag="12"></SELECT></TD>
							</TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>세액계산여부</TD>
				                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txttax_calc" TAG="2X" VALUE="Y" CHECKED ID="tax_calc1"><LABEL FOR="tax_calc1">Y</LABEL>&nbsp;
				                                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txttax_calc" TAG="2X" VALUE="N" ID="tax_calc2"><LABEL FOR="tax_calc2">N</LABEL>
				                                     <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txttax_calc" TAG="2X" VALUE="H" ID="tax_calc3"><LABEL FOR="tax_calc3">기존세액유지</LABEL></TD>

						    </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>급여구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_cd" ALT="급여구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH=13 SIZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmp()">&nbsp;<INPUT NAME="txtName" TYPE="Text" MAXLENGTH=30 SIZE=20 tag=14XXXU></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



