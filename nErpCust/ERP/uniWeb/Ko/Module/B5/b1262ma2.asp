<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : B1262MA2                    
'*  4. Program Name         : 매출거래처형태등록                   
'*  5. Program Desc         : 매출거래처형태등록                
'*  6. Comproxy List        : PB5GS42.dll, PB5GS43.dll     
'*  7. Modified date(First) : 2001/01/05               
'*  8. Modified date(Last)  : 2001/12/18                
'*  9. Modifier (First)     : Kim Hyungsuk                
'* 10. Modifier (Last)      : Sonbumyeol 
'* 11. Comment              :                
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 2002/12/02 : INCLUDE 추가 성능 적용, Kang Jun Gu
'*                            2002/12/09 : INCLUDE 다시 성능 적용, Kang Jun Gu
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit       
                                                      '☜: indicates that All variables must be declared in advance
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "b1262mb2.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "b1262ma8"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop      ' Popup
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

 frm1.rdoParttype1.checked = True
 frm1.rdoParttype21.checked = True
 frm1.rdoUsage_flag1.checked = True
 frm1.txtBp_cd1.focus

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Function OpenBp_cd(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
 Case 0
  arrParam(0) = "주문처"      <%' 팝업 명칭 %>

  arrParam(2) = Trim(frm1.txtBp_cd1.Value)		<%' Code Condition%>
  arrParam(3) = ""								<%' Name Cindition%>

  arrParam(5) = "주문처"					<%' TextBox 명칭 %>

  arrHeader(0) = "주문처"					<%' Header명(0)%>
  arrHeader(1) = "주문처명"					<%' Header명(1)%>
  frm1.txtBp_cd1.focus 
 Case 1
  If frm1.txtBp_cd2.readOnly = True Then
   IsOpenPop = False
   Exit Function
  End If

  arrParam(0) = "주문처"      <%' 팝업 명칭 %>
  
  arrParam(2) = Trim(frm1.txtBp_cd2.Value)		<%' Code Condition%>
  arrParam(3) = ""								<%' Name Cindition%>

  arrParam(5) = "주문처"					<%' TextBox 명칭 %>

  arrHeader(0) = "주문처"					<%' Header명(0)%>
  arrHeader(1) = "주문처명"					<%' Header명(1)%>
  frm1.txtBp_cd2.focus 
 Case 2 
  arrParam(0) = "파트너거래처"    <%' 팝업 명칭 %>
   
  arrParam(2) = Trim(frm1.txtPartner_cd1.Value) <%' Code Condition%>
  arrParam(3) = ""								<%' Name Cindition%>

  arrParam(5) = "파트너거래처"				<%' TextBox 명칭 %>

  arrHeader(0) = "파트너거래처"				<%' Header명(0)%>
  arrHeader(1) = "파트너거래처명"			<%' Header명(1)%>
  frm1.txtPartner_cd1.focus 
 Case 3
  If frm1.txtPartner_cd2.readOnly = True Then
   IsOpenPop = False
   Exit Function
  End If

  arrParam(0) = "파트너거래처"    <%' 팝업 명칭 %>
  
  arrParam(2) = Trim(frm1.txtPartner_cd2.Value) <%' Code Condition%>
  arrParam(3) = ""								<%' Name Cindition%>

  arrParam(5) = "파트너거래처"				<%' TextBox 명칭 %>

  arrHeader(0) = "파트너거래처"				<%' Header명(0)%>
  arrHeader(1) = "파트너거래처명"			<%' Header명(1)%>
  frm1.txtPartner_cd2.focus 
 End Select
 
 arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE 명칭 %>

 arrParam(4) = "BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & " )"			<%' Where Condition%>
   
 arrField(0) = "BP_CD"							<%' Field명(0)%>
 arrField(1) = "BP_NM"							<%' Field명(1)%>
    
  
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetBp_cd(arrRet, iWhere)
 End If 
 
End Function

'========================================================================================================= 
Function SetBp_cd(Byval arrRet, Byval iWhere)

 With frm1

  Select Case iWhere
  Case 0
   .txtBp_cd1.value = arrRet(0) 
   .txtBp_nm1.value = arrRet(1)   
  Case 1
   .txtBp_cd2.value = arrRet(0) 
   .txtBp_nm2.value = arrRet(1)   
   lgBlnFlgChgValue = True
  Case 2
   .txtPartner_cd1.value = arrRet(0) 
   .txtPartner_nm1.value = arrRet(1)   
  Case 3
   .txtPartner_cd2.value = arrRet(0) 
   .txtPartner_nm2.value = arrRet(1)   
   lgBlnFlgChgValue = True
  End Select

 End With
 
End Function

'========================================================================================================= 
Function CookiePage(ByVal Kubun)
 
 On Error Resume Next

 Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>

 Dim strTemp, arrVal

 If Kubun = 1 Then

  WriteCookie CookieSplit , frm1.txtBp_cd1.value  & parent.gRowSep & frm1.txtBp_nm1.value

 ElseIf Kubun = 0 Then

  strTemp = ReadCookie(CookieSplit)
   
  If strTemp = "" Then Exit Function
   
  arrVal = Split(strTemp, parent.gRowSep)

  If arrVal(0) = "" Then Exit Function
  
  frm1.txtBp_cd1.value =  arrVal(0)
  frm1.txtBp_nm1.value =  arrVal(1)
  frm1.txtPartner_cd1.value = arrVal(2)
  frm1.txtPartner_nm1.value = arrVal(3)
  
  select case arrVal(4)
  case "SSH"
   frm1.rdoParttype1.checked = True
  case "SBI"
   frm1.rdoParttype2.checked = True
  case "SPA"
   frm1.rdoParttype3.checked = True
  case else
   frm1.rdoParttype1.checked = True
  end select

  if Err.number <> 0 then
   Err.Clear
   WriteCookie CookieSplit , ""
   exit function
  end if
  
  Call MainQuery()  
   
  WriteCookie CookieSplit , ""
  
 End If
 
End Function

'========================================================================================================= 
Function JumpChgCheck()

 Dim IntRetCD

 '************ 싱글인 경우 **************
 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then Exit Function
 End If

 Call CookiePage(1)
 Call PgmJump(BIZ_PGM_JUMP_ID)

End Function


'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call SetDefaultVal
	Call InitVariables              '⊙: Initializes local global variables

	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------

	'폴더/조회/입력 
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전 
	'/다음/복사/엑셀 
	'/인쇄/찾기 
	Call SetToolBar("11101000000011")          '⊙: 버튼 툴바 제어 
	Call CookiePage(0)

End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'========================================================================================================= 
Sub rdoParttype21_OnClick()
 lgBlnFlgChgValue = True
End Sub

Sub rdoParttype22_OnClick()
 lgBlnFlgChgValue = True
End Sub

Sub rdoParttype23_OnClick()
 lgBlnFlgChgValue = True 
End Sub

Sub rdoUsage_flag1_OnClick()
 lgBlnFlgChgValue = True
End Sub

Sub rdoUsage_flag2_OnClick()
 lgBlnFlgChgValue = True
End Sub

Sub chkPartner_OnClick()
 lgBlnFlgChgValue = True 
End Sub

'========================================================================================================= 
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>
    
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If


<%  '-----------------------
    'Check RadioButton area
    '-----------------------%>
 If frm1.rdoParttype1.checked = True Then 
  frm1.txtRadioType.value = "SSH"
 ElseIf frm1.rdoParttype2.checked = True Then
  frm1.txtRadioType.value = "SBI"
 ElseIf frm1.rdoParttype3.checked = True Then
  frm1.txtRadioType.value = "SPA"
 End IF

 Call ggoOper.LockField(Document, "N")                                        <%'⊙: Lock  Suitable  Field%>
 Call SetToolBar("11101000000011")

      
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery                <%'☜: Query db data%>
       
    FncQuery = True                <%'⊙: Processing is OK%>
        
End Function

'========================================================================================================= 
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x") 
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'⊙: Clear Condition,Contents Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄 

    Call SetToolBar("11101000000011")
    Call SetDefaultVal
    
    FncNew = True                <%'⊙: Processing is OK%>

End Function

'========================================================================================================= 
Function FncDelete() 
    
    dim IntRetCD
    
    FncDelete = False              <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If

    <% 'SINGLE일경우만 해당 %>
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
<%  '-----------------------
    'Check RadioButton area
    '-----------------------%>
 If frm1.rdoParttype1.checked = True Then 
  frm1.txtRadioType.value = "SSH"
 ElseIf frm1.rdoParttype2.checked = True Then
  frm1.txtRadioType.value = "SBI"
 ElseIf frm1.rdoParttype3.checked = True Then
  frm1.txtRadioType.value = "SPA"
 End IF

<%  '-----------------------
    'Delete function call area
    '-----------------------%>
    Call DbDelete               <%'☜: Delete db data%>
    
    FncDelete = True                                                        <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then                             <%'⊙: Check contents area%>
       Exit Function
    End If


<%  '-----------------------
    'Check CheckBox area
    '-----------------------%>

 If frm1.chkPartner.checked = True Then
  frm1.txtCheck.value = "Y"
 Else
  frm1.txtCheck.value = "N"
 End If 

<%  '-----------------------
    'Check RadioButton area
    '-----------------------%>
 If frm1.rdoParttype21.checked = True Then 
  frm1.txtRadioType.value = "SSH"
 ElseIf frm1.rdoParttype22.checked = True Then
  frm1.txtRadioType.value = "SBI"
 ElseIf frm1.rdoParttype23.checked = True Then
  frm1.txtRadioType.value = "SPA"
 End IF


 If frm1.rdoUsage_flag1.checked = True Then
  frm1.txtRadioFlag.value = "Y" 
 ElseIf frm1.rdoUsage_flag2.checked = True Then
  frm1.txtRadioFlag.value = "N" 
 End IF
    
<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncCopy() 
 Dim IntRetCD

    If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
'    lgIntFlgMode = parent.OPMD_CMODE            <%'⊙: Indicates that current mode is Crate mode%>
    
    <% ' 조건부 필드를 삭제한다. %>
    Call ggoOper.ClearField(Document, "1")                                      <%'⊙: Clear Condition Field%>
    Call ggoOper.LockField(Document, "N")         <%'⊙: This function lock the suitable field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>
    Call SetToolBar("11101000000011")

	frm1.rdoParttype1.checked = True
    
    frm1.txtBp_cd2.focus
    lgBlnFlgChgValue = True
End Function

'========================================================================================================= 
Function FncPrint() 
 Call Parent.FncPrint()
End Function

'========================================================================================================= 
Function FncExcel() 
 Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================= 
Function FncFind() 
 Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================= 
Function FncExit()
 Dim IntRetCD
 FncExit = False
    If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
  'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
  If IntRetCD = vbNo Then
   Exit Function
  End If
    End If
    FncExit = True
End Function

'========================================================================================================= 
Function DbDelete() 
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbDelete = False              <%'⊙: Processing is NG%>
    
        
	If   LayerShowHide(1) = False Then
        Exit Function 
    End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003       <%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&txtBp_cd2=" & Trim(frm1.txtBp_cd2.value)   <%'☜: 삭제 조건 데이타 %>
    strVal = strVal & "&txtPartner_cd2=" & Trim(frm1.txtPartner_cd2.value) <%'☜: 삭제 조건 데이타 %>
    strVal = strVal & "&txtRadioType=" & Trim(frm1.txtRadioType.value)  <%'☜: 삭제 조건 데이타 %>    
    
	Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbDelete = True                                                         <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbDeleteOk()              <%'☆: 삭제 성공후 실행 로직 %>
 Call FncNew()
End Function

'========================================================================================================= 
Function DbQuery() 
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
        
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       <%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&txtBp_cd1=" & Trim(frm1.txtBp_cd1.value)   <%'☜: 조회 조건 데이타 %>
    strVal = strVal & "&txtPartner_cd1=" & Trim(frm1.txtPartner_cd1.value) <%'☜: 조회 조건 데이타 %>
    strVal = strVal & "&txtRadioType=" & Trim(frm1.txtRadioType.value)  <%'☜: 조회 조건 데이타 %>    
 Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True                                                          <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'⊙: This function lock the suitable field%>
 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
 Call SetToolBar("11111000001111")

    frm1.txtBp_nm1.value = frm1.txtBp_nm2.value 
    frm1.txtPartner_nm1.value = frm1.txtPartner_nm2.value 

End Function

'========================================================================================================= 
Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>

 DbSave = False               <%'⊙: Processing is NG%>

     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

 
    Dim strVal

 With frm1
  .txtMode.value = parent.UID_M0002           <%'☜: 비지니스 처리 ASP 의 상태 %>
  .txtFlgMode.value = lgIntFlgMode
  .txtInsrtUserId.value = parent.gUsrID 
  .txtUpdtUserId.value = parent.gUsrID

  Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
 
 End With
 
    DbSave = True                                                           <%'⊙: Processing is NG%>
    
End Function

'========================================================================================================= 
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>

    frm1.txtBp_cd1.value = frm1.txtBp_cd2.value 
    frm1.txtPartner_cd1.value = frm1.txtPartner_cd2.value 

	If frm1.txtRadioType.value = "SSH" Then
	 frm1.rdoParttype1.checked = True
	ElseIf frm1.txtRadioType.value = "SBI" Then
	 frm1.rdoParttype2.checked = True
	ElseIf frm1.txtRadioType.value = "SPA" Then
	 frm1.rdoParttype3.checked = True
	End If
            
    Call InitVariables
    
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR >
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출거래처형태등록</font></td>
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
     <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>주문처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT NAME="txtBp_cd1" ALT="주문처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd frm1.txtBp_cd1.value,0">&nbsp;
             <INPUT NAME="txtBp_nm1" TYPE="Text" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>파트너거래처</TD>
         <TD CLASS="TD6" NOWRAP><INPUT NAME="txtPartner_cd1" ALT="파트너거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPartner_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd frm1.txtPartner_cd1.value,2">&nbsp;
             <INPUT NAME="txtPartner_nm1" TYPE="Text" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>거래처형태</TD>
         <TD CLASS="TD6" NOWRAP>
          <input type=radio CLASS="RADIO" name="rdoParttype" id="rdoParttype1" value="SSH" tag = "11" checked>
           <label for="rdoParttype1">납품처</label>
          <input type=radio CLASS="RADIO" name="rdoParttype" id="rdoParttype2" value="SBI" tag = "11">
           <label for="rdoParttype2">발행처</label>
          <input type=radio CLASS = "RADIO" name="rdoParttype" id="rdoParttype3" value="SPA" tag = "11">
           <label for="rdoParttype3">수금처</label></TD>
         <TD CLASS=TD5 NOWRAP></TD>
         <TD CLASS=TD6 NOWRAP></TD>
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD WIDTH=100% VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS=TD5 NOWRAP>주문처</TD>
        <TD CLASS=TD656 NOWRAP><INPUT NAME="txtBp_cd2" ALT="주문처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd frm1.txtBp_cd2.value,1">&nbsp;
             <INPUT NAME="txtBp_nm2" TYPE="Text" SIZE=25 tag="24"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>거래처형태</TD>       
        <TD CLASS=TD656 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoParttype_2" id="rdoParttype21" value="SSH" tag = "23" checked>
          <label for="rdoParttype21">납품처</label>
         <input type=radio CLASS="RADIO" name="rdoParttype_2" id="rdoParttype22" value="SBI" tag = "23">
          <label for="rdoParttype22">발행처</label>
         <input type=radio CLASS = "RADIO" name="rdoParttype_2" id="rdoParttype23" value="SPA" tag = "23">
          <label for="rdoParttype23">수금처</label></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>파트너거래처</TD>
        <TD CLASS=TD656 NOWRAP><INPUT NAME="txtPartner_cd2" ALT="파트너거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPartner_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd frm1.txtPartner_cd2.value,3">&nbsp;
             <INPUT NAME="txtPartner_nm2" TYPE="Text" SIZE=25 tag="24"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>사용여부</TD>
        <TD CLASS=TD656 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="Y" tag = "21" checked>
          <label for="rdoUsage_flag1">예</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="N" tag = "21">
          <label for="rdoUsage_flag2">아니오</label></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>거래처담당자명</TD>
        <TD CLASS=TD656 NOWRAP><INPUT NAME="txtBp_prsn_nm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="21"></TD>
       </TR>
       <TR>       
        <TD CLASS=TD5 NOWRAP>거래처담당자연락처</TD>
        <TD CLASS=TD656 NOWRAP><INPUT NAME="txtBp_contact_pt" TYPE="Text" MAXLENGTH="30" SIZE=40 tag="21"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP></TD>
        <TD CLASS=TD656 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkPartner" tag="21XXX" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkPartner">주파트너 거래처</LABEL></TD>
       </TR>
       <%Call SubFillRemBodyTD656(8)%>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR HEIGHT=20>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck()">거래처형태조회</a></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
