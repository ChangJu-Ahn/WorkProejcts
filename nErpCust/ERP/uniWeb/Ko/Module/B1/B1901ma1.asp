<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Common Format
'*  3. Program ID           :
'*  4. Program Name         : 공통포맷등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/02
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young/Lee Seok Gon
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             <% '☜: indicates that All variables must be declared in advance %>

Const BIZ_PGM_ID  = "b1901mb1.asp"											<% '☆: 비지니스 로직 ASP명 %>
Const BIZ_PGM_COUNT_FORMAT = "b1902ma1"
Const BIZ_PGM_NUMERIC_FORMAT = "b1903ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

Dim lgMpsFirmDate, lgLlcGivenDt											<% '☜: 비지니스 로직 ASP에서 참조하므로 %>

Dim lgCurName()															<%'☆ : 개별 화면당 필요한 로칼 전역 변수 %>
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2

Sub InitVariables()

    lgIntFlgMode = "C"                                               <%'⊙: Indicates that current mode is Create mode%>
    lgBlnFlgChgValue = False                                                <%'⊙: Indicates that no value changed%>
    lgIntGrpCount = 0                                                       <%'⊙: Initializes Group View Size%>
    <%'----------  Coding part  -------------------------------------------------------------%>
    IsOpenPop = False														<%'☆: 사용자 변수 초기화 %>
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub


Sub InitComboBox()	
	'B0011 => 날짜포맷	
	Call CommonQueryRs(" MINOR_NM,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0011", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboDateFormat, lgF0, lgF1, Chr(11))
	
	'B0012 => 소수포맷 
	Call CommonQueryRs(" MINOR_CD,MINOR_CD ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0012", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboDecimalCharacter, lgF0, lgF1, Chr(11))
	
End Sub

Sub Form_Load()
																		'⊙: Load Common DLL
    Call InitVariables																'⊙: Initializes local global variables    
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------           
    Call InitComboBox
    Call SetToolbar("1100100000001111") 
    
    Call FncQuery  ''DBQuery
    
	frm1.cboDateFormat.focus	
End Sub

Function LoadCountFormat()
    
    PgmJump(BIZ_PGM_COUNT_FORMAT)

End Function

Function LoadNumericFormat()
    
    PgmJump(BIZ_PGM_NUMERIC_FORMAT)

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")										<%'⊙: Clear Contents  Field%>
    Call InitVariables															<%'⊙: Initializes local global variables%>
    
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = false Then														<%'☜: Query db data%>
		Exit Function 
	End If	
	
	FncQuery = True																<%'⊙: Processing is OK%>        
End Function

function FncSave()     
    Dim intCntUser
    Dim IntRetCD
    
    On Error Resume Next
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    frm1.txtLogInCnt.value = "0"	'현재로그인 user를 초기화 시킴 
    
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    End If

    If frm1.cboDateFormat.value = "" Then
		Call DisplayMsgBox("970021", "X", "날짜포맷", "X")
		Exit Function
    ElseIf frm1.cboDecimalCharacter.value = "" Then
		Call DisplayMsgBox("970021", "X", "소수점 구분자", "X")
		Exit Function
    End If
    
    If Not chkField(Document, "2") Then                             <%'⊙: Check contents area%>
       Exit Function
    End If
    
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)                                                   <%'☜: Protect system from crashing%>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbDeleteOk()														<%'☆: 삭제 성공후 실행 로직 %>
	Call FncNew()
End Function

Function DbQuery() 
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
    
    Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
	
    DbQuery = True                                                          <%'⊙: Processing is NG%>
End Function

Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = "U"												<%'⊙: Indicates that current mode is Update mode%>
    
End Function

Function DbSave() 
    Err.Clear																<%'☜: Protect system from crashing%>

	DbSave = False															<%'⊙: Processing is NG%>

    Dim strVal
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = parent.UID_M0002											<%'☜: 비지니스 처리 ASP 의 상태 %>
		.txtFlgMode.value = lgIntFlgMode
		.txtDate.value = frm1.cboDateFormat.value
		.txtDecimal.value = frm1.cboDecimalCharacter.value
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	End With

    DbSave = True       
End Function

Function DbSaveOk()															<%'☆: 저장 성공후 실행 로직 %>
    Call InitVariables
    
    Call MainQuery()

End Function

Function CheckLogInUser() 
    Dim IntRetCD 
	Dim strLogInCnt
	Dim arrRet
	Dim arrParam(5)
    Dim tempMsg
    Dim iCalledAspName
    
	arrParam(0) = ""
	arrParam(1) = ""
    
    Err.Clear			
    strLogInCnt = Cint(frm1.txtLogInCnt.value)
    
    tempMsg = "접속중인 사용자가 존재하므로 저장할 수 없습니다 " & vbCrLf
    tempMsg = tempMsg & "이 자료는 시스템관리자 1명만 접속했을 때 저장할 수 있습니다" & vbCrLf
    tempMsg = tempMsg & "접속중인 사용자 정보를 보시겠습니까?"
      
    intRetCD = MsgBox(tempMsg,vbExclamation + vbYesNo, Parent.gLogoName & "-[Warning]")
    
    If IntRetCD = vbNo Then
		Exit Function
	End If

	iCalledAspName = AskPRAspName("LoginUserList")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "LoginUserList", "X")
		lgIsOpenPop = False
		Exit Function
	End If


	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),, "dialogWidth=400px; dialogHeight=600px; center: Yes; help: No; resizable: No; status: No;")


End Function


Function cboDateFormat_onChange()
	lgBlnFlgChgValue = True
End Function

Function cboDecimalCharacter_onChange()
	lgBlnFlgChgValue = True
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공통포맷</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=40>
					<TD CLASS="TD5"></TD>
					<TD CLASS="TD6"></TD>
				</TR>
				<TR HEIGHT=20>		
					<TD CLASS="TD5">날짜포맷</TD>
					<TD CLASS="TD6"><SELECT NAME="cboDateFormat" tag="22X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>									
				</TR>
				<TR HEIGHT=20>
					<TD CLASS="TD5">소수점 구분자</TD>
					<TD CLASS="TD6"><SELECT NAME="cboDecimalCharacter" tag="22X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>
				</TR>
				<TR HEIGHT=40>
					<TD CLASS="TD5"></TD>
					<TD CLASS="TD6"></TD>				
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadCountFormat">수량포맷</A>&nbsp;|&nbsp;<A HREF="vbscript:LoadNumericFormat">Numeric포맷</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="b1901mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtDate" tag="24"><INPUT TYPE=HIDDEN NAME="txtDecimal" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLogInCnt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
