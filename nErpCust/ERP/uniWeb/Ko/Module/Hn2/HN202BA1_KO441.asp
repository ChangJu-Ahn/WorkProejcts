<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 급여계산 
*  3. Program ID           : H6009ba1
*  4. Program Name         : H6009ba1
*  5. Program Desc         : 급여관리/급여계산 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/07
*  8. Modified date(Last)  : 2003/06/13
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
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "HN202BB1_KO441.asp"
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
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
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
	
    'Call ggoOper.FormatDate(frm1.txtBas_yy, Parent.gDateFormat, 3)

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtBas_yy.Year = strYear 
	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "BA") %>
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
   
   'If pOpt = "Q" Then
   '   lgKeyStream = Frm1.txtWarrentNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   'Else
   '   lgKeyStream = Frm1.txtMajorCd.Value & Parent.gColSep         'You Must append one character(Parent.gColSep)
   'End If   

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
    
    'Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0045", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProv_cd, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
		
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtBas_yy, Parent.gDateFormat, 3)                         '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
	Call InitVariables                                                     '⊙: Setup the Spread sheet
	
	Call InitComboBox()

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
    frm1.txtBas_yy.focus
			
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

      
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
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
	Call Parent.FncExport(Parent.C_SINGLE)
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
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
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

		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub


'======================================================================================================
' Function Name : ExeSend
' Function Desc : 
'=======================================================================================================
Function ExeSend()
	 
	Dim strVal
	Dim IntRetCD
    Dim strSQL
	Dim strGubun

	Dim strBasYY, strProvCd

	ExeSend = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	strGubun  = "S"
	strSQL		 = ""
	strBasYY = Trim(frm1.txtBas_yy.value)
	strProvCd = Trim(frm1.txtProv_cd.value)
'	strYYYYMMDD = UniConvDateAToB(frm1.txtBas_yy,Parent.gDateFormatYYYYMM,Parent.gServerDateFormat)

 '   IF  FuncAuthority("!", strYYYYMMDD, Parent.gUsrID) = "N" THEN
 '       Call DisplayMsgBox("800294","X","X","X")
 '       Call BtnDisabled(0)
 '       exit function
 '   END IF

	strBasYY = frm1.txtBas_yy.Year
    'StrSQL = StrSQL2 & " AND (pay_cd IS NULL OR tax_cd IS NULL)"
    'StrSQL = StrSQL & " AND (retire_dt IS NULL OR retire_dt > " & UNIConvDate(strPreDate) & ")"
    'StrSQL = StrSQL & " AND entr_dt <=  " & FilterVar(UNIConvDate(frm1.txtBas_yy.text), "''", "S") & ""
	StrSQL = StrSQL & " EVAL_YY = " &  FilterVar(strBasYY, "''", "S") 
	If Not IsNull(strProvCd) And strProvCd <> "" Then StrSQL = StrSQL & " AND EVAL_TYPE = " &  FilterVar(strProvCd, "''", "S") 

    If 	CommonQueryRs(" COUNT(*) "," HBA040T ", strSQL, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then

        IF Trim(Replace(lgF0, Chr(11), "")) <> 0 THEN        
			IntRetCD = DisplayMsgBox("800397",Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				'Call BtnDisabled(0)
				Exit Function
			Else
				strGubun = "S"				
			End If
        END IF
    End if

	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	Call BtnDisabled(1)
	

	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	'strVal = strVal & "&txtBas_yy=" & Replace(strPayYYMMDt, Parent.gComDateType, "")
	strVal = strVal & "&strGubun=" & strGubun
	strVal = strVal & "&txtBas_yy=" & strBasYY
	strVal = strVal & "&txtProv_Cd=" & strProvCd

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeSend = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeSendOk
' Function Desc : ExeSend가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeSendOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function

Function ExeSendNo()				            '☆: 실행된 자료가 없습니다 
	Dim IntRetCD 

    Call DisplayMsgBox("800161","X","X","X")

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


' Function Name : ExeDelete
' Function Desc : 
'=======================================================================================================
Function ExeDelete()
	 
	Dim strVal
	Dim IntRetCD
    Dim strSQL
	Dim strGubun

	Dim strBasYY, strProvCd

	ExeDelete = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	strGubun  = "S"
	strSQL		 = ""
	strBasYY = Trim(frm1.txtBas_yy.value)
	strProvCd = Trim(frm1.txtProv_cd.value)

	strBasYY = frm1.txtBas_yy.Year

	StrSQL = StrSQL & " EVAL_YY = " &  FilterVar(strBasYY, "''", "S") 
	If Not IsNull(strProvCd) And strProvCd <> "" Then StrSQL = StrSQL & " AND EVAL_TYPE = " &  FilterVar(strProvCd, "''", "S") 

    If 	CommonQueryRs(" COUNT(*) "," HBA040T ", strSQL, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then

        IF Trim(Replace(lgF0, Chr(11), "")) <> 0 THEN        
			IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			Else
				strGubun = "D"				
			End If
        END IF
    End if

	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	Call BtnDisabled(1)

	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	'strVal = strVal & "&txtBas_yy=" & Replace(strPayYYMMDt, Parent.gComDateType, "")
	strVal = strVal & "&strGubun=" & strGubun
	strVal = strVal & "&txtBas_yy=" & strBasYY
	strVal = strVal & "&txtProv_Cd=" & strProvCd

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeDelete = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function


Sub txtBas_yy_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBas_yy.Action = 7
		frm1.txtBas_yy.focus
	End If
End Sub

Sub txtBas_yy_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>평가결과수신</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>평가년도</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtBas_yy NAME="txtBas_yy" CLASS=FPDTYYYY  title=FPDATETIME ALT="기준년" tag="12X1" VIEWASTEXT> </OBJECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>평가구분</TD>
			            	    <TD CLASS="TD6"><SELECT NAME="txtProv_cd" CLASS ="cbonormal" tag="11" ALT="평가구분"><OPTION Value=""></OPTION></SELECT></TD>
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
					<TD><BUTTON NAME="btnSend"  CLASS="CLSSBTN" onclick="ExeSend()" Flag=1>반영</BUTTON>&nbsp;
							<BUTTON NAME="btnExe2" CLASS="CLSSBTN" onclick="ExeDelete()" Flag=1>취소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



