<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 기본급테이블인상 
*  3. Program ID           : h6001ba1
*  4. Program Name         : h6001ba1
*  5. Program Desc         : 기본급테이블인상 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/19
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
Const BIZ_PGM_ID = "h6001bb1.asp"
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
    
    lgStrPrevKey = ""                           'initializes Previous Key
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
	
    frm1.txtpost_apply_dt.focus()
	frm1.txtpost_apply_dt.text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)

	Call CommonQueryRs(" MAX(APPLY_STRT_DT)","HDF010T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtapply_dt.text =  UNIConvDateDBToCompany(Replace(lgF0,Chr(11),""),Null)
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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   Dim strYear, strMonth, strDay

    With frm1
        lgKeyStream = UniConvDateToYYYYMMDD(.txtapply_dt.Text, Parent.gDateFormat, "") & Parent.gColSep
        lgKeyStream = lgKeyStream     & UniConvDateToYYYYMMDD(.txtpost_apply_dt.Text, Parent.gDateFormat, "") & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtPay_grd.value  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtPay_grd_nm.value  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtpay_grd1.value  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtpay_grd2.value  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtallow_cd.value  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & .txtallow_amt.text & Parent.gColSep     '7
	    lgKeyStream = lgKeyStream     & .txtallow_rate.text & Parent.gColSep    '8
	    lgKeyStream = lgKeyStream     & .txtbase_amt.value & Parent.gColSep      '9
	    lgKeyStream = lgKeyStream     & .txtend_type.value  & Parent.gColSep
    End With    
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr    
    Dim iCodeArr1 
    Dim iNameArr1    

	Call CommonQueryRs("ALLOW_NM,ALLOW_CD","HDA010T","HDA010T.CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND HDA010T.ALLOW_KIND = " & FilterVar("1", "''", "S") & "  ORDER BY HDA010T.ALLOW_CD ASC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr  = lgF0
    iCodeArr  = lgF1
    Call SetCombo2(frm1.txtallow_cd,iCodeArr, iNameArr,Chr(11))            ''''''''DB에서 불러 condition에서    
    
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0052", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1  = lgF0
    iCodeArr1  = lgF1
    Call SetCombo2(frm1.txtend_type,iCodeArr1, iNameArr1,Chr(11))            ''''''''DB에서 불러 condition에서       
End Sub

'========================================================================================================
' Name : CookiePage
' Desc : 기본급 테이블 등록 페이지에서 jump 할경우 기존기준일을 기본급테이블등록 페이지 값을 가져온다.
'========================================================================================================
Function CookiePage(ByVal flgs)
    Dim strReadCookie
    strReadCookie = ReadCookie("APPLY_DT")

    If  strReadCookie<>"" Then
        frm1.txtapply_dt.Text = strReadCookie
	    WriteCookie "APPLY_DT" , ""
    	FncQuery()
    End If
End Function
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format	
	Call AppendNumberPlace("6","3","2")
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet

	Call InitComboBox()
	frm1.txtpost_apply_dt.focus
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어			
	Call CookiePage(0)
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
    Err.Clear                                                                    '☜: Clear err status
	FncQuery = false
    if  txtPay_grd_onChange()    then
        Exit Function
    End If
	FncQuery = true
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
	Call FncNew()	
End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect(iWhere)
	Call BtnDisabled(1)
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	Err.Clear 

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If   
    
   	If Len(Trim(frm1.txtapply_dt.Text)) And Len(Trim(frm1.txtpost_apply_dt.Text)) Then

		If UniConvDateToYYYYMMDD(frm1.txtapply_dt.Text,Parent.gDateFormat,"") >= UniConvDateToYYYYMMDD(frm1.txtpost_apply_dt.Text,Parent.gDateFormat,"") Then
			Call DisplayMsgBox("800443","X", frm1.txtpost_apply_dt.Alt, frm1.txtapply_dt.Alt)   
			frm1.txtpost_apply_dt.Focus
            Set gActiveElement = document.activeElement                            
			Call BtnDisabled(0)
			Exit Function
		End If

	End If

	If Not FncQuery() Then
	    Call BtnDisabled(0)
	    Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    Call MakeKeyStream("X")
    
    Select Case iWhere      
        Case 1                          '==>실행버튼 클릭 
	        strVal = BIZ_PGM_ID & "?txtMode="     & Parent.UID_M0002
	    Case 2                          '==>삭제버튼 클릭 
	        strVal = BIZ_PGM_ID & "?txtMode="     & Parent.UID_M0003
	End Select
	strVal = strVal     & "&lgKeyStream=" & lgKeyStream

	Call RunMyBizASP(MyBizASP, strVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
	Dim IntRetCD 

    Call DisplayMsgBox("800161","X","X","X")

End Function

Sub txtapply_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtapply_dt.Action = 7
		frm1.txtapply_dt.focus
	End If
End Sub
Sub txtpost_apply_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtpost_apply_dt.Action = 7
		frm1.txtpost_apply_dt.focus
	End If
End Sub
'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	        arrParam(0) = "급호코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtPay_grd.value       			' Code Condition
	    	arrParam(3) = ""'frm1.txtPay_grd_nm.value 				' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "급호코드" 		    	        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "급호코드"	   		        	' Header명(0)
	    	arrHeader(1) = "급호코드명"	    		        ' Header명(1)
	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPay_grd.focus	
		Exit Function
	Else
		Call SetCode(arrRet)       	
	End If	

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet)
     frm1.txtPay_grd.value = arrRet(0)       			' Code Condition
	 frm1.txtPay_grd_nm.value = arrRet(1) 				' Name Cindition
	 frm1.txtPay_grd.focus
End Function
'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function txtPay_grd_OnChange()
    Dim IntRetCd

    If Trim(frm1.txtPay_grd.value) = "" Then
        frm1.txtPay_grd_nm.Value = ""
        
    Else
        IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtPay_grd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtPay_grd.Value)<>""  Then
            frm1.txtPay_grd_nm.Value=""
            Call DisplayMsgBox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다.
			txtPay_grd_OnChange = true
        Else
            frm1.txtPay_grd_nm.Value=Trim(Replace(lgF0,Chr(11),""))
        
        End If
    End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기본급테이블인상</font></td>
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
									<TD CLASS=TD5 NOWRAP>기존기준일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h6001ba1_txtapply_dt_txtapply_dt.js'></script></TD>
								</TR>
                                <TR>
									<TD CLASS=TD5 NOWRAP>수정기준일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h6001ba1_txtpost_apply_dt_txtpost_apply_dt.js'></script></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5">급호</TD>
								    <TD CLASS=TD6 NOWRAP>
								        <INPUT NAME="txtPay_grd"     SIZE=10  MAXLENGTH=2  ALT ="직급" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCode()">
								        <INPUT NAME="txtPay_grd_nm"  SIZE=20  MAXLENGTH=50 TAG="14XXXU"></TD>
								</TR>			

							    <TR>
								    <TD CLASS="TD5">호봉</TD>
								    <TD CLASS=TD6 NOWRAP>
								        <INPUT NAME="txtpay_grd1" SIZE=10  MAXLENGTH=3 ALT ="시작호봉" TAG="11XXXU">부터
								        <INPUT NAME="txtpay_grd2" SIZE=10  MAXLENGTH=3 ALT ="종료호봉" TAG="11XXXU">까지</TD>
								</TR>			

							    <TR>
									<TD CLASS=TD5 NOWRAP>기본급수당</TD>
				                    <TD CLASS=TD6 NOWRAP><SELECT NAME="txtallow_cd" ALT="기본급수당" CLASS ="cbonormal" TAG="12XN"></SELECT></TD>
				                    
							    </TR>
								<TR>
									<TD CLASS="TD5">정액</TD>
								    <TD CLASS=TD6 NOWRAP>
								        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
								        <TR><TD><script language =javascript src='./js/h6001ba1_txtallow_amt_txtallow_amt.js'></script>정율</TD>
								            <TD><script language =javascript src='./js/h6001ba1_txtallow_rate_txtallow_rate.js'></script>% </TD>
								        </TR>
								        </TABLE></TD>								        
							    </TR>							    
								<TR>
									<TD CLASS="TD5">처리기준</TD>
								    <TD CLASS=TD6 NOWRAP>
								        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
								        <TR><TD NOWRAP><script language =javascript src='./js/h6001ba1_txtbase_amt_txtbase_amt.js'></script>원 미만</TD>
								            <TD NOWRAP><SELECT NAME="txtend_type" ALT="처리타입" CLASS ="cbonormal" TAG="11XN"><OPTION></OPTION></SELECT></TD>
								        </TR>
								        </TABLE></TD>								        
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onClick="VBScript: Call ExeReflect(1)" Flag=1>실행</BUTTON>&nbsp;
					    <BUTTON NAME="btnDelete" CLASS="CLSSBTN" onClick = "VBScript: Call ExeReflect(2)" Flag=1>삭제</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

