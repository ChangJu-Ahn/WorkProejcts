<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Yim Young Ju
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

On Error Resume Next
	Err.Clear  

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Sub SetDB()
    If frm1.SDBC.checked = true Then
       frm1.sdb.value ="unierp272"
    Else
       frm1.sdb.value =""
    End If   
End Sub

Sub doWork()
	
	Dim i

    If Trim(frm1.SDB.value) = "" Then
       'MsgBox "대상 데이터 베이스를 선택하십시요."
       Call DisplayMsgBox("990050","X","X","X")
       frm1.SDB.focus
       Exit Sub
    End If   

    If Trim(frm1.TDB.value) = "" Then
       'MsgBox "생성할 데이터 베이스명을 입력하십시요." 
       Call DisplayMsgBox("990051","X","X","X")
       frm1.TDB.focus
       Exit Sub
    End If   

    If UCASE(Trim(frm1.SDB.value)) = UCASE(Trim(frm1.TDB.value))  Then
       'MsgBox "원본과 타켓 데이터베이스가 같습니다." 
       Call DisplayMsgBox("990052","X","X","X") 
       frm1.TDB.focus
       Exit Sub
    End If   
    
    If frm1.TDBC.checked = False Then
    
		For i=1 to frm1.SDB.length - 1 '990058
			IF UCASE(Trim(frm1.TDB.value)) = UCASE(frm1.SDB(i).value) THEN
				'MSGBOX "이미 존재하는 데이터베이스입니다."
				Call DisplayMsgBox("990058","X","X","X") 
				EXIT SUB
			END IF
		Next
    
    End If    
    

    If frm1.METHOD_A.checked = True Then
       frm1.action = "ZK000MA3.asp"
    Else
       frm1.action = "ZK000MA3_1.asp"
    End If   
       frm1.submit
       
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<%

Call LoadBasisGlobalInf()

MasterConnString = MakeConnString(GetGlobalInf("gDBServerIP") ,GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),"master")      


%>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1  METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0 >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>컴퍼니생성</font></td>
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
								<TD CLASS="TD5" NOWRAP>생성방식</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=RADIO NAME=METHOD ID=METHOD_A VALUE = A CLASS=RADIO CHECKED>컴퍼니 복사<INPUT TYPE=RADIO NAME=METHOD ID=METHOD_B VALUE = B CLASS=RADIO> 신규컴퍼니
								</OBJECT>
								</TD>															
							</TR>
							<TR>
						  		<TD CLASS=TD5 NOWRAP>Source</TD>
								<TD CLASS=TD6 NOWRAP>
								<SELECT NAME=SDB><%=DBList(MasterConnString)%></SELECT>
								</TD>
						        </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>Target</TD>
				        	    <TD CLASS="TD6">
				        	    <INPUT  TYPE=TEXT tag="21XXXU" MAXLENGTH="20" NAME=TDB>
				        	    <INPUT TYPE=CHECKBOX id=TDBC name=TDBC CLASS=RADIO                       >기존 DB에 강제생성
				        	    </TD>
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:doWork()" Flag=1>작업시작</BUTTON>
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

</BODY>
</HTML>


