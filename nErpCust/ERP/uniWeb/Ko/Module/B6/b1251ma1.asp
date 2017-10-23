<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1251ma1
'*  4. Program Name         : 구매그룹등록 
'*  5. Program Desc         : 구매그룹등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
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
<!--'******************************************  1.1 Inc 선언   **********************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->    
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css"> 

<!--'==========================================  1.1.2 공통 Include   ======================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             <% '☜: indicates that All variables must be declared in advance %>

'==========================================  1.2.1 Global 상수 선언  ======================================

Const BIZ_PGM_ID  = "b1251mb1.asp"            '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID  = "b1251qa1"

'==========================================  1.2.2 Global 변수 선언  =====================================
Dim lgBlnFlgChgValue    '☜: Variable is for Dirty flag
Dim lgIntGrpCount    '☜: Group View Size를 조사할 변수 
Dim lgIntFlgMode    '☜: Variable is for Operation Status

Dim lgNextNo     '☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo     ' ""

Dim IsOpenPop          

'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub


Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then

		strTemp = ReadCookie("GroupCd")
		  
		If strTemp = "" then Exit Function
		  
		frm1.txtGroupCd1.value = ReadCookie("GroupCd")
		frm1.txtGroupNm1.value = ReadCookie("GroupNm")
		WriteCookie "GroupCd" , ""
		WriteCookie "GroupNm" , ""
		WriteCookie "CallType" , ""
		  
		MainQuery()
	Else 
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	     
	  WriteCookie "Kubun" , "Y"
	  WriteCookie "OrgCd", Trim(frm1.txtOrgCd2.value)
	  WriteCookie "GroupCd", Trim(frm1.txtGroupCd1.value)
	     
	  Call PgmJump(BIZ_PGM_JUMP_ID)
	End IF
End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                              '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                      '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                             '⊙: Initializes Group View Size

    IsOpenPop = False						'☆: 사용자 변수 초기화 
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtGroupCd1.focus
	Set gActiveElement = document.activeElement
	Call SetToolbar("1110100000001111")          '⊙: 버튼 툴바 제어 
End Sub

'------------------------------------------  OpenGroup1()  -------------------------------------------------
Function OpenGroup1(byval strComp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd1.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹" 
	arrParam(1) = "B_Pur_Grp"
 
	if strComp="Group1" then
		 arrParam(2) = UCase(Trim(frm1.txtGroupCd1.Value))
	else
		 arrParam(2) = UCase(Trim(frm1.txtGroupCd2.Value))
	end if 
 
	arrParam(4) = ""
	arrParam(5) = "구매그룹"   
 
	arrField(0) = "PUR_GRP" 
	arrField(1) = "PUR_GRP_NM" 
	   
	arrHeader(0) = "구매그룹"  
	arrHeader(1) = "구매그룹명"
	   
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		If strComp="Group1" Then
			frm1.txtGroupCd1.focus
		Else
			frm1.txtGroupCd2.focus
		End If
		Exit Function
	Else
		If strComp="Group1" Then
			frm1.txtGroupCd1.Value= arrRet(0)  
			frm1.txtGroupNm1.Value= arrRet(1)  
			frm1.txtGroupCd1.focus
		Else
			frm1.txtGroupCd2.Value= arrRet(0)  
			frm1.txtGroupNm2.Value= arrRet(1)  
			frm1.txtGroupCd2.focus
			lgBlnFlgChgValue = True
		End If 
	End If 
 
End Function

'------------------------------------------  OpenCost()  -------------------------------------------------
Function OpenCost()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "비용집계처"      ' 팝업 명칭 
	arrParam(1) = "B_COST_CENTER"       ' TABLE 명칭 
	 
	arrParam(2) = UCase(Trim(frm1.txtCostCd.Value))  ' Code Condition
	 
	arrParam(4) = ""       ' Where Condition
	arrParam(5) = "비용집계처"       ' TextBox 명칭 
	 
	arrField(0) = "COST_CD"     ' Field명(0)
	arrField(1) = "COST_NM"     ' Field명(1)
	    
	arrHeader(0) = "비용집계처"      ' Header명(0)
	arrHeader(1) = "비용집계처명"    ' Header명(1)
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		frm1.txtCostCd.Value    = arrRet(0)  
		frm1.txtCostNm.Value    = arrRet(1)  
		lgBlnFlgChgValue = True
		frm1.txtCostCd.focus
	End If 
End Function

'------------------------------------------  OpenORG1()  -------------------------------------------------
Function OpenORG1(byval strcomp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직"      <%' 팝업 명칭 %>
	arrParam(1) = "B_Pur_Org"      <%' TABLE 명칭 %>
	 
	arrParam(2) = UCase(Trim(frm1.txtORGCd2.Value)) <%' Code Condition%>
	 
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "       <%' Where Condition%>
	arrParam(5) = "구매조직"       <%' TextBox 명칭 %>
	 
	arrField(0) = "PUR_ORG"     <%' Field명(0)%>
	arrField(1) = "PUR_ORG_NM"     <%' Field명(1)%>
	    
	arrHeader(0) = "구매조직"      <%' Header명(0)%>
	arrHeader(1) = "구매조직명"      <%' Header명(1)%>
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtORGCd2.focus
		Exit Function
	Else
		frm1.txtOrgCd2.Value= arrRet(0)  
		frm1.txtOrgNm2.Value= arrRet(1)  
		lgBlnFlgChgValue = True
		frm1.txtORGCd2.focus
	End If 
End Function

'Radio에서 Click을 할 경우 flag를 Setting
Sub Setchangeflg()
	lgBlnFlgChgValue = True 
End Sub

'사용자가 Radio Button을 Click할 때 마다 숨겨진 txtUseflg를 Setting 
Sub Changeflg()
	If frm1.rdoUseflg(0).checked = true Then
	 frm1.txtUseflg.value= "Y"
	Else
	 frm1.txtUseflg.value= "N"
	End If 
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call loadInfTB19029
    Call ggoOper.LockField(Document, "N") 
         '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables   
   
    '----------  Coding part  -------------------------------------------------------------
    
    Call SetDefaultVal
    Call Changeflg
	Call CookiePage(0)
End Sub

'==========================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrDt.Focus
	End if
End Sub

'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToDt.Focus
	End if
End Sub

'==========================================================================================
Sub txtFrDt_Change()
	lgBlnFlgChgValue = true 
End Sub

'==========================================================================================
Sub txtToDt_Change()
	lgBlnFlgChgValue = true 
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
  If lgBlnFlgChgValue = True Then
	IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
  End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")          '⊙: Clear Contents  Field
    Call InitVariables               '⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then         '⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function
    Call Changeflg
    FncQuery = True                '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables               '⊙: Initializes local global variables
    
    Call SetDefaultVal
    
    FncNew = True                '⊙: Processing is OK
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then Exit Function

    FncDelete = False  
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                     
		Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                                      
    
End Function

'========================================================================================
 Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not ChkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    Call ChangeFlg
    
    with frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then   
		   Call DisplayMsgBox("17a003","X","유효일","X")   
			Exit Function
		End if   
	End with
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE            '⊙: Indicates that current mode is Crate mode
    
    ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")         '⊙: This function lock the suitable field
    Call Changeflg
    
    frm1.txtGroupCd2.value = ""
    frm1.txtGroupNm2.value = ""
    
    lgBlnFlgChgValue = True
End Function

'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

'========================================================================================
Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                             
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001      
    strVal = strVal & "&txtGroupCd1=" & lgPrevNo       '☆: 조회 조건 데이타 
    
    Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       '☜: 비지니스 처리 ASP의 상태값 
    strVal = strVal & "&txtGroupCd1=" & lgNextNo       '☆: 조회 조건 데이타 
    
 Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)            '☜: 화면 유형 
End Function

'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False 
                 '⊙: Processing is NG
    if LayerShowHide(1) = False then
       Exit Function 
    end if
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003       '☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtGroupCd1=" & Trim(frm1.txtGroupCd2.value)  '☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 
 
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'========================================================================================
Function DbDeleteOk()              '☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function

'========================================================================================
 Function DbQuery() 
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery = False                                                         '⊙: Processing is NG
    
    if LayerShowHide(1) = False then
       Exit Function 
    end if
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       '☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtGroupCd1=" & Trim(frm1.txtGroupCd1.value)

	Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG

End Function

'========================================================================================
Function DbQueryOk()              '☆: 조회 성공후 실행로직 
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE            '⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")         '⊙: This function lock the suitable field
	lgBlnFlgChgValue = False

    Call SetToolbar("11111000001111")
    frm1.txtGroupNm2.focus
End Function

'========================================================================================
Function DbSave() 
    Err.Clear                '☜: Protect system from crashing

	DbSave = False               '⊙: Processing is NG

    If LayerShowHide(1) = False Then
       Exit Function 
    End If
    
    Dim strVal

	With frm1
	 .txtMode.value = Parent.UID_M0002           '☜: 비지니스 처리 ASP 의 상태 
	 .txtFlgMode.value = lgIntFlgMode
	 
	   Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
	 
	End With

	DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()               '☆: 저장 성공후 실행 로직 
    frm1.txtGroupCd1.value = frm1.txtGroupCd2.value 
    frm1.txtGroupNm1.value = frm1.txtGroupNm2.value 

    Call InitVariables
    Call MainQuery()              
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR>
  <TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=500>&nbsp;</TD>
     <TD WIDTH=*>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR HEIGHT=*>
  <TD WIDTH=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>> 
    <TR>
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100% valign=top>
      <FIELDSET CLASS="CLSFLD">
      <TABLE <%=LR_SPACE_TYPE_40%>>
       <TR>
        <TD CLASS="TD5" NOWRAP>구매그룹</TD>
        <TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd1" SIZE=10 MAXLENGTH=4  tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup1('Group1')">
                <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm1" NAME="arrCond" tag="14X"></TD>
       </TR>
      </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD WIDTH=100% valign=top> 
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS="TD5" NOWRAP>구매그룹</TD>
        <TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd2" MAXLENGTH=4 SIZE=10  tag="23XXXU">&nbsp;&nbsp;&nbsp;&nbsp;
                <INPUT TYPE=TEXT Alt="구매그룹명" NAME="txtGroupNm2" MAXLENGTH=50 SIZE=40 tag="22" ></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>구매조직</A></TD>
        <TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT Alt="구매조직" MAXLENGTH=4 NAME="txtORGCd2" SIZE=10 MAXLENGTH=4  tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG1('ORG2')">
                <INPUT TYPE=TEXT Alt="구매조직명" NAME="txtORGNm2" SIZE=20 tag="24x"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>비용집계처</TD>
        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" ALT="비용집계처" SIZE=10 MAXLENGTH=10 tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCost()">
                <INPUT TYPE=TEXT NAME="txtCostNm" ALT="비용집계처" SIZE=20 tag="24X" ></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>사용여부</TD>
        <TD CLASS="TD656" NOWRAP><INPUT TYPE=radio Alt="사용여부" ID="rdoUseflgY" NAME="rdoUseflg" class="radio" tag="2X" Value="Y" checked id="rdoUseflgY" ONCLICK="vbscript:SetChangeflg()"><label for="rdoUseflgY">&nbsp;예&nbsp;</label>
                <INPUT TYPE=radio Alt="사용여부" ID="rdoUseflgN" NAME="rdoUseflg" class="radio" tag="2X" Value="N" id="rdoUseflgN" ONCLICK="vbscript:SetChangeflg()"><label for="rdoUseflgN">&nbsp;아니오&nbsp;</label></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>유효일</TD>
        <TD CLASS="TD656" NOWRAP>
         <table cellspacing=0 cellpadding=0>
          <tr>
           <td>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효일 NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </td>
           <td>~</td>
           <td>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효일 NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="21" Title="FPDATETIME"></OBJECT>');</SCRIPT>
           </td>
          <tr>
         </table>
        </TD>
       </TR>
		<%Call SubFillRemBodyTD656(17)%>
      </TABLE>
     </TD> 
    </TR>
   </TABLE>
  </TD>
 </TR>    
 <tr>
  <td HEIGHT="3"></td>
    </tr>
    <tr HEIGHT="20">
  <td WIDTH="100%">
   <table WIDTH="100%">
    <tr>
     <td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">구매그룹조회</a></td>
     <td WIDTH="10"></td>
    </tr>
   </table>
  </td>
    </tr>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUseflg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOrgCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOrgNm" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>


