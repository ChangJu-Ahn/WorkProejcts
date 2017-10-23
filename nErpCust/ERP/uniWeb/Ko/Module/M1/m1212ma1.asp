<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1212MA1
'*  4. Program Name         : 공급처칼렌다생성 
'*  5. Program Desc         : 공급처칼렌다생성 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/01/16
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop   

<!--'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************!-->
Const BIZ_PGM_ID = "m1212mb1.asp"
Const BIZ_PGM_CHANGE_CAL = "m1212ma2"

Const C_Month = 1          
Const C_Day = 2
Const C_Remark = 3

<!-- '------------------------------------------  OpenBpCd()  -------------------------------------------------
' Name : OpenBpCd()
' Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"      
	arrParam(1) = "B_Biz_Partner"     
	arrParam(2) = Trim(frm1.txtBpCd.Value)   
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"      
	 
	arrField(0) = "BP_CD"       
	arrField(1) = "BP_NM"       
	    
	arrHeader(0) = "공급처"      
	arrHeader(1) = "공급처명"     
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If 
End Function

Function LoadChangeCal()
    parent.WriteCookie "Bp_Cd",UCase(Trim(frm1.txtBpCd.value))
    parent.WriteCookie "Year",frm1.cboYear.value
    PgmJump(BIZ_PGM_CHANGE_CAL)
End Function
<!-- '==========================================  2.2.1 InitComboBox()  ========================================
' Name : InitComboBox()
' Description : ComboBox 초기화 
'========================================================================================================= !-->
Sub InitComboBox()
	Dim i, ii
	For i=<%=Year(GetSvrDate)%>-10 To <%=Year(GetSvrDate)%>+10
		Call SetCombo(frm1.cboYear, i, i)
	Next
	    
	frm1.cboYear.value = <%=Year(GetSvrDate)%>
End Sub
<!-- '******************************************  3.1 Window 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'********************************************************************************************************* !-->
<!-- '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= !-->
Sub Form_Load()
    Call ggoOper.LockField(Document, "N")                   
    
    Call SetToolbar("1000000000001111")      
    Call InitComboBox     
	frm1.txtBpCd.focus 
	Set gActiveElement = document.activeElement
End Sub

<!--
'========================================================================================
' Function Name : GenOk
' Function Desc : GenOk가 성공적일 경우 MyBizASP 에서 호출되는 Function, 
'========================================================================================
-->
Function GenOk()
	Call DisplayMsgBox("183114","X","X","X")
End Function

<!--
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
-->
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False) 
End Function
<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
Function FncPrint() 
 Call parent.FncPrint()
End Function

<!--
'========================================================================================
' Function Name : btnBatch_OnClick
' Function Desc : 칼렌다 생성 버튼을 누르면 칼렌다를 생성한다.
'========================================================================================
-->
Function btnBatch_OnClick()
	Dim strVal,IntRetCD
	 
	If Not chkField(Document, "1") Then   
		Exit Function
	End If
	  
	IntRetCD = DisplayMsgBox("17a006", parent.VB_YES_NO,"칼렌다 생성", "X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
	 
	Call chkCheckBox()
	  
	With frm1
		.txtInsrtUserId.value = parent.gUsrID    
		If LayerShowHide(1) = False then
			Exit Function 
		End if  
		       
		Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	End With
End Function

Function chkCheckBox()
	If frm1.chkSun.checked = true Then
		frm1.chkSun.value = "Y"
	Else
		frm1.chkSun.value = "N"
	End If
	 
	If frm1.chkMon.checked = true Then
		frm1.chkMon.value = "Y"
	Else
		frm1.chkMon.value = "N"
	End If 
	 
	If frm1.chkTue.checked = true Then
		frm1.chkTue.value = "Y"
	Else
		frm1.chkTue.value = "N"
	End If

	If frm1.chkWed.checked = true Then
		frm1.chkWed.value = "Y"
	Else
		frm1.chkWed.value = "N"
	End If

	If frm1.chkThu.checked = true Then
		frm1.chkThu.value = "Y"
	Else
		frm1.chkThu.value = "N"
	End If

	If frm1.chkFri.checked = true Then
		frm1.chkFri.value = "Y"
	Else
		frm1.chkFri.value = "N"
	End If

	If frm1.chkSat.checked = true Then
		frm1.chkSat.value = "Y"
	Else
		frm1.chkSat.value = "N"
	End If
 
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
 Function FncExit() 
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 

</HEAD>

<!-- '#########################################################################################################
'            6. Tag부 
'######################################################################################################### !-->
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
     <TD>
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공급처칼렌다</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
       </TR>
      </TABLE>
     </TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD WIDTH=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
       <TR>
     <TD WIDTH=100%>
      <TABLE <%=LR_SPACE_TYPE_60%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>공급처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd" MAXLENGTH=10 SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
                 <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>   
        </TR>
        <TR>
         <TD CLASS="TD5" CLASS="TD5">생성년도</TD>
         <TD CLASS="TD6" WIDTH=20>
          <SELECT Name="cboYear" ALT=생성년도 tag="22" CLASS=cboNormal></SELECT>
         </TD>         
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD WIDTH=100% valign=top>
      <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <TD COLSPAN=1 ALIGN=left>휴무일지정</TD>
       </TR>       
       <TR>
        <TD COLSPAN=1>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkSun" checked><label for="chkSun">일</label>
            <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkMon"><label for="chkMon">월</label>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkTue"><label for="chkTue">화</label>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkWed"><label for="chkWed">수</label>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkThu"><label for="chkThu">목</label>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkFri"><label for="chkFri">금</label>
         <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkSat"><label for="chkSat">토</label>
        </TD>
       </TR>
       <TR>
        <TD COLSPAN=2 ALIGN=left><HR></TD>
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
     <TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" >칼렌다 생성</BUTTON></TD>
     <TD WIDTH=* ALIGN=RIGHT> <A href="vbscript:LoadChangeCal">칼렌다수정</TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
