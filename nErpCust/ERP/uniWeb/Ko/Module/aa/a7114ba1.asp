<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--*******************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Asset Management
'*  3. Program ID           : a7114ba1.asp
'*  4. Program Name         : 감가상각 결과 반영 
'*  5. Program Desc         :
'*  6. Comproxy List        : AS0071 
'                             
'                             
'*  7. Modified date(First) : 2000/12/31
'*  8. Modified date(Last)  : 2000/12/31
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!--'화면처리ASP에서 서버작업이 필요한 경우  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!--: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 
 '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "a7114bb1.asp"  
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 

 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '******************************************  2.1 Pop-Up 함수   **********************************************
'	기능: Pop-Up 
'********************************************************************************************************* 

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Data Code PopUp
'--------------------------------------------------------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function fnButtonExec()
    Dim strVal       
    Dim strWorkDt
	Dim RetFlag
	Dim strYear
	Dim strMonth
	Dim strDay

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then        '⊙: Check contents area
       Exit Function
    End If
            
	 RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
	''RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")
	If RetFlag = VBNO Then
		Exit Function
	End IF

    Err.Clear    	
    Call LayerShowHide(1) 


    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
	    
    if frm1.Rb_WK1.checked = true then
		strVal = strVal & "&txtRadio=" & "1"								'☆: 조회 조건 데이타 
    else
		strVal = strVal & "&txtRadio=" & "2"								'☆: 조회 조건 데이타 
	end if
    
    
    
    Call ExtractDateFrom(frm1.fpDateTime1.Text,frm1.fpDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	strWorkDt = strYear & strMonth
     
    strVal = strVal & "&txtWKyymm=" & strWorkDt
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
end function

Function fnButtonExecOk()
    Dim IntRetCD 

	IntRetCD = DisplayMsgBox("990000","X","X","X")   '☜ 바뀐부분	
End function

 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
Function FncQuery()
End Function
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function
Function FncPrint() 
	Parent.fncPrint()    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True    
End Function

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "BA") %>
End Sub

 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029	 
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)       
   
    Call ggoOper.FormatDate(frm1.txtWKyymm, gDateFormat, 2)	
    Call ggoOper.LockField(Document, "N")	
    Call SetToolbar("10000000000011")

    frm1.fpDateTime1.Text =  UNIMonthClientFormat(parent.gFiscEnd)
    
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'=======================================================================================================
'   Event Name : txtWKyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtWKyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtWKyymm.Action = 7
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>> <%' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감가상각결과반영</font></td>
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
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업 구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>결과 반영</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2><LABEL FOR=Rb_WK2>반영 취소</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업기준년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/a7114ba1_fpDateTime1_txtWKyymm.js'></script>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btn배치" CLASS="CLSMBTN" OnClick="VBScript:Call fnButtonExec()" Flag=1>실 행</BUTTON> &nbsp</TD>		        
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>

		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

