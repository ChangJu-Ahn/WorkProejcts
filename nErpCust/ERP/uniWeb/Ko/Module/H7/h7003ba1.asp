<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 개인상여조정율일괄생성 
*  3. Program ID           : h6001ba1
*  4. Program Name         : h6001ba1
*  5. Program Desc         : 개인상여조정율일괄생성 
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
Const BIZ_PGM_ID = "h7003bb1.asp"
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
	
	frm1.txtBonus_yymm.focus
	frm1.txtbonus_yymm.text = UniConvDateAToB("<%=GetsvrDate%>", Parent.gServerDateFormat, Parent.gDateFormatYYYYMM)
	frm1.txtbonus_yymm1.text = frm1.txtbonus_yymm.text
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim BonusDt, BonusDt1
	
    BonusDt		= frm1.txtBonus_yymm.Year & Right("0" & frm1.txtBonus_yymm.month, 2)    
    BonusDt1	= frm1.txtBonus_yymm1.Year & Right("0" & frm1.txtBonus_yymm1.month, 2)    
   
    With frm1
	    lgKeyStream =					Trim(BonusDt) & Parent.gColSep
	    lgKeyStream = lgKeyStream     & Trim(.txtBonus_type.value) & Parent.gColSep
	    lgKeyStream = lgKeyStream     & Trim(BonusDt1)  & Parent.gColSep
	    lgKeyStream = lgKeyStream     & Trim(.txtBonus_type1.value)  & Parent.gColSep
    End With    

End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr    

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " And Minor_cd between " & FilterVar("2", "''", "S") & " and " & FilterVar("9", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtBonus_type,iCodeArr, iNameArr,Chr(11))

    Call SetCombo2(frm1.txtBonus_type1,iCodeArr, iNameArr,Chr(11))
End Sub
'========================================================================================================
' Name : CookiePage
' Desc : 기본급 테이블 등록 페이지에서 jump 할경우 기존기준일을 기본급테이블등록 페이지 값을 가져온다.
'========================================================================================================
Function CookiePage(ByVal flgs)
    If  ReadCookie("BONUS_YYMM_DT")<>"" Then
        frm1.txtBonus_yymm.Text  = ReadCookie("BONUS_YYMM_DT")
        frm1.txtBonus_type.Value = ReadCookie("BONUS_TYPE")
	    WriteCookie "BONUS_TYPE" , ""
	    WriteCookie "BONUS_YYMM_DT" , ""
    End If
End Function
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format	
	Call AppendNumberPlace("6","4","2")

	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtBonus_yymm, Parent.gDateFormat, 2)    
	Call ggoOper.FormatDate(frm1.txtBonus_yymm1, Parent.gDateFormat, 2)

	Call InitVariables                                                     '⊙: Setup the Spread sheet

	Call InitComboBox()
	frm1.txtBonus_yymm.focus()
	
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
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD
	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	Err.Clear 

	If Not chkField(Document, "1") Then
		Exit Function
	End If
    
  	If (Trim(frm1.txtBonus_yymm.Text)=Trim(frm1.txtBonus_yymm1.Text)) And (frm1.txtBonus_type.value = frm1.txtBonus_type1.value) Then
			Call DisplayMsgBox("800271","X","X","X")   '처리대상 데이터가 동일합니다.
			frm1.txtBonus_yymm.Focus
            Set gActiveElement = document.activeElement                            
			Exit Function
	End If
    If Not(ValidDateCheck(frm1.txtBonus_yymm, frm1.txtBonus_yymm1)) Then
        frm1.txtBonus_yymm.text = ""
        frm1.txtBonus_yymm1.Text = ""
        frm1.txtBonus_yymm.focus
        Exit Function
    End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	If LayerShowHide(1) = False then
    		Exit Function 
    End if
    
    MakeKeyStream("X")
	strVal = BIZ_PGM_ID & "?txtMode="     & Parent.UID_M0002
	strVal = strVal     & "&lgKeyStream=" & lgKeyStream

	Call RunMyBizASP(MyBizASP, strVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG

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

Sub txtBonus_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBonus_yymm.Action = 7
		frm1.txtBonus_yymm.focus
	End If
End Sub
Sub txtBonus_yymm1_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBonus_yymm1.Action = 7
		frm1.txtBonus_yymm1.focus
	End If
End Sub
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>개인상여조정율일괄생성</font></td>
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
			                	<TD CLASS="TD5" NOWRAP>대상 상여년월</TD>
			                	<TD CLASS="TD6" NOWRAP>
			                		<script language =javascript src='./js/h7003ba1_fpDateTime2_txtBonus_yymm.js'></script>
			                	</TD>
			                </TR>
			                <TR>
			                	<TD CLASS="TD5" NOWRAP>대상 상여구분</TD>
			                	<TD CLASS="TD6" NOWRAP><SELECT NAME="txtBonus_type" ALT="처리 상여구분" STYLE="WIDTH:150px" TAG="12"></SELECT></TD>
			                </TR>
			                <TR>
			                	<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
			                	<TD CLASS="TD6" NOWRAP></TD>
			                </TR>
			                <TR>			
			                	<TD CLASS="TD5" NOWRAP>처리 상여년월</TD>
			                	<TD CLASS="TD6" NOWRAP>
			                		<script language =javascript src='./js/h7003ba1_fpDateTime2_txtBonus_yymm1.js'></script>
			                	</TD>
			                </TR>
			                <TR>
			                	<TD CLASS="TD5" NOWRAP>처리 상여구분</TD>
			                	<TD CLASS="TD6" NOWRAP><SELECT NAME="txtBonus_type1" ALT="처리 상여구분" STYLE="WIDTH:150px" TAG="12"></SELECT></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect(1)" Flag="1">실행</BUTTON>&nbsp;
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


