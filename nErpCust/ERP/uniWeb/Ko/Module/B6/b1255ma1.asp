<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : B1255MA1
'*  4. Program Name         : 영업조직등록 
'*  5. Program Desc         : 영업조직등록 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2001/12/21
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Park insik
'* 11. Comment              :
'* 12. Comment              : 2002/12/02 : INCLUDE 추가 성능 적용, Kang Jun Gu
'* 13. Comment              : 2002/12/09 : INCLUDE 다시 성능 적용, Kang Jun Gu
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                       

<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID = "b1255mb1.asp"            
Const BIZ_PGM_JUMP_ID = "b1255ma8"
Dim IsOpenPop   

'===============================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
End Sub

'===============================================================================================================
Sub SetDefaultVal()
	With frm1
	 .txtSales_Org1.focus  
	 .txtRadio.value = "Y"
	 .txtEndOrgFlag.value = "Y" 
	 .rdoUsage_flag1.checked = True
	 .rdoEndOrgFlagY.checked = True
	End With
End Sub

'===============================================================================================================
Function OpenSorgCode()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

	If Len(frm1.txtlvl.value) = 0 Then
		Call DisplayMsgBox("205152", "x", frm1.txtlvl.Alt, "x")
		frm1.txtlvl.focus
		Exit Function
	End If

 IsOpenPop = True

 arrParam(0) = "상위영업조직"     

 arrParam(1) = "B_SALES_ORG"       
 arrParam(2) = Trim(frm1.txtUpper_Sales_Org.value) 
 arrParam(4) = "END_ORG_FLAG = " & FilterVar("N", "''", "S") & "  and LVL < " & Trim(frm1.txtlvl.value) 
 arrParam(5) = "상위영업조직"     
 
 arrField(0) = "SALES_ORG"       
 arrField(1) = "SALES_ORG_NM"      
    
 arrHeader(0) = "상위영업조직"      
 arrHeader(1) = "상위영업조직명"      
   
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetSorgCode(arrRet)
 End If 
 
End Function
'===============================================================================================================
Function OpenConSorgCode()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "영업조직"     
 arrParam(1) = "B_SALES_ORG"      
 arrParam(2) = Trim(frm1.txtSales_Org1.value) 
 arrParam(3) = "" 
 arrParam(4) = ""        
 arrParam(5) = "영업조직"     
 
    arrField(0) = "SALES_ORG"      
    arrField(1) = "SALES_ORG_NM"     
    
    arrHeader(0) = "영업조직"     
    arrHeader(1) = "영업조직명"     
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetConSorgCode(arrRet)
 End If 
End Function

'===============================================================================================================
Function SetSorgCode(Byval arrRet)

	If arrRet(0) <> "" Then 
	 frm1.txtUpper_Sales_Org.value = arrRet(0)
	 frm1.txtUpper_Sales_OrgNm.value = arrRet(1)
	 lgBlnFlgChgValue = True
	 frm1.txtUpper_Sales_Org.focus
	End IF
 
End Function

'===============================================================================================================
Function SetConSorgCode(byval arrRet)

	If arrRet(0) <> "" Then 
	 frm1.txtSales_Org1.value = arrRet(0)  
	 frm1.txtSales_Org_nm1.value = arrRet(1)
	 frm1.txtSales_Org1.focus  
	End If

End Function
'===============================================================================================================
Sub CookiePage(Byval Kubun)

 On Error Resume Next

 Const CookieSplit = 4877      

 Dim strTemp, arrVal

 If Kubun = 1 Then

  WriteCookie CookieSplit , frm1.txtSales_Org1.value & parent.gRowSep & frm1.txtSales_Org_nm1.value _
      & parent.gRowSep & frm1.txtUpper_Sales_Org.value & parent.gRowSep & frm1.txtUpper_Sales_OrgNm.value _
      & parent.gRowSep & frm1.txtlvl.value & parent.gRowSep & frm1.txtRadio.value
                  
 ElseIf Kubun = 0 Then

  strTemp = ReadCookie(CookieSplit)

  If strTemp = "" then Exit Sub
  
  arrVal = Split(strTemp, parent.gRowSep)

  If arrVal(0) = "" then Exit Sub  

  frm1.txtSales_Org1.value =  arrVal(0)
  frm1.txtSales_Org_nm1.value =  arrVal(1)
  
  If Err.number <> 0 then
   Err.Clear
   WriteCookie CookieSplit , ""
   Exit Sub
  End If
  
  Call MainQuery()
  
  WriteCookie CookieSplit , ""

 End IF
 
End Sub


'===============================================================================================================
Function JumpChgCheck()

 Dim IntRetCD


 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
  If IntRetCD = vbNo Then Exit Function
 End If

 Call CookiePage(1)
 Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'===============================================================================================================
Function NumericCheck()

 Dim objEl, KeyCode
 
 Set objEl = window.event.srcElement
 KeyCode = window.event.keycode

 Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
 Case Else
  window.event.keycode = 0
 End Select

End Function


'================================================================================================================
Sub Form_Load()

    Call InitVariables              
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) <% '⊙: Format Contents  Field %>
    Call ggoOper.LockField(Document, "N")                                   
	Call SetDefaultVal
    
    Call SetToolBar("11101000000011")          
	Call CookiePage(0)

End Sub


'================================================================================================================
Sub rdoUsage_flag1_OnClick()
 lgBlnFlgChgValue = True
 frm1.txtRadio.value = "Y" 
End Sub

Sub rdoUsage_flag2_OnClick()
 lgBlnFlgChgValue = True
 frm1.txtRadio.value = "N" 
End Sub

Sub rdoEndOrgFlagY_OnClick()
 lgBlnFlgChgValue = True
 frm1.txtEndOrgFlag.value = "Y" 
End Sub

Sub rdoEndOrgFlagN_OnClick()
 lgBlnFlgChgValue = True
 frm1.txtEndOrgFlag.value = "N" 
End Sub

'================================================================================================================
Sub txtlvl_onKeyPress()
 Call NumericCheck()
End Sub

'================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")          
    Call InitVariables              
    
    If Not chkField(Document, "1") Then          
       Exit Function
    End If

	Call ggoOper.LockField(Document, "N")        
    Call SetToolBar("11101000000011")          
    
    Call DbQuery                

    FncQuery = True                
        
End Function

'=============================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")   
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                          
    Call ggoOper.LockField(Document, "N")                                       
    Call InitVariables              

    Call SetToolBar("11101000000011")          
    Call SetDefaultVal
    
    FncNew = True                

End Function

'=============================================================================================================
Function FncDelete() 

	Dim IntRetCD
	     
	FncDelete = False              
	    
	If lgIntFlgMode <> parent.OPMD_UMODE Then      
	    Call DisplayMsgBox("900002","x","x","x")
	    Exit Function
	End If
	 

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")

	If IntRetCD = vbNo Then Exit Function
	    

	Call DbDelete    
	    
	FncDelete = True                                                        
    
End Function

'=============================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "2") Then                             
       Exit Function
    End If

	If frm1.rdoUsage_flag1.checked = True Then
	 frm1.txtRadio.value = frm1.rdoUsage_flag1.value
	Else
	 frm1.txtRadio.value = frm1.rdoUsage_flag2.value
	End If

	If frm1.rdoEndOrgFlagY.checked = True Then
	 frm1.txtEndOrgFlag.value = frm1.rdoEndOrgFlagY.value
	Else
	 frm1.txtEndOrgFlag.value = frm1.rdoEndOrgFlagN.value
	End If

    CAll DbSave                                                    
    
    FncSave = True                                                          
    
End Function

'================================================================================================================
Function FncCopy() 
 Dim IntRetCD

    If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
  
  If IntRetCD = vbNo Then
      Exit Function
  End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE            
        
    Call ggoOper.ClearField(Document, "1")                                  
    Call ggoOper.LockField(Document, "N")         
    Call InitVariables              <%%>
    Call SetToolBar("11101000000011")
    
    frm1.txtSales_Org2.value = ""
    frm1.txtSales_Org_nm2.value = ""
    frm1.txtSales_Org2.focus
    
    lgBlnFlgChgValue = True
    
End Function

Function FncCancel() 
    On Error Resume Next                                                    
End Function


Function FncInsertRow() 
     On Error Resume Next                                                   
End Function


Function FncDeleteRow() 
    On Error Resume Next                                                    
End Function


Function FncPrint() 
 Call Parent.FncPrint()
End Function

'===============================================================================================================
Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002","x","x","x")  '☜ 바뀐부분 
        
        Exit Function 
    End If
 
	If   LayerShowHide(1) = False Then
	    Exit Function 
	End If

	 
	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
    strVal = strVal & "&txtSales_Org2=" & Trim(frm1.txtSales_Org2.value) <%'☆: 조회 조건 데이타 %>
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  <%'☆: 조회 조건 데이타 %>
         
 Call RunMyBizASP(MyBizASP, strVal)

End Function

'===============================================================================================================
Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","x","x","x") 
        Exit Function
    End If

  
	If   LayerShowHide(1) = False Then
	    Exit Function 
	End If

	 
	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
    strVal = strVal & "&txtSales_Org2=" & Trim(frm1.txtSales_Org2.value) 
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)  
    
 Call RunMyBizASP(MyBizASP, strVal)

End Function

'===============================================================================================================
Function FncExcel() 
 Call Parent.FncExport(parent.C_SINGLE)
End Function


Function FncFind() 
 Call Parent.FncFind(parent.C_SINGLE, False)
End Function


Function FncExit()
 Dim IntRetCD
 FncExit = False
    If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
  If IntRetCD = vbNo Then
   Exit Function
  End If
    End If
    FncExit = True
End Function

'================================================================================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False              
     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003       
    strVal = strVal & "&txtSales_Org2=" & Trim(frm1.txtSales_Org2.value) <%'☜: 삭제 조건 데이타 %>
    
 Call RunMyBizASP(MyBizASP, strVal)          
 
    DbDelete = True                                                         

End Function

'================================================================================================================
Function DbDeleteOk()              
 lgBlnFlgChgValue = False
 Call MainNew()
End Function

'================================================================================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         
  
 
	If   LayerShowHide(1) = False Then
	    Exit Function 
	End If

    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
    strVal = strVal & "&txtSales_Org2=" & Trim(frm1.txtSales_Org1.value)  <%'☆: 조회 조건 데이타 %>

 Call RunMyBizASP(MyBizASP, strVal)          
 
    DbQuery = True                                                          

End Function
'================================================================================================================
Function DbQueryOk()              
 
    lgIntFlgMode = parent.OPMD_UMODE            
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         

	Call SetToolBar("1111100011111111")
 
	frm1.txtSales_Org_nm1.value = frm1.txtSales_Org_nm2.value 
	frm1.txtSales_Org1.focus
 
End Function

'================================================================================================================
Function DbSave() 

    Err.Clear                

	DbSave = False               

  
	If   LayerShowHide(1) = False Then
	    Exit Function 
	End If

 
    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002           
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID 
		.txtUpdtUserId.value = parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)         
	 
	End With
 
    DbSave = True                                                           
    
End Function

'================================================================================================================
Function DbSaveOk()               

    frm1.txtSales_Org1.value = frm1.txtSales_Org2.value 
    
    Call InitVariables
    
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</head>
<body TABINDEX="-1" SCROLL="no">
<form NAME="frm1" TARGET="MyBizASP" METHOD="POST">
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color="white"><%=Request("strASPMnuMnuNm")%></font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
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
         <TD CLASS="TD5" NOWRAP>영업조직</TD>
         <TD CLASS="TD656" NOWRAP><input NAME="txtSales_Org1" TYPE="Text" MAXLENGTH="4" tag="12XXXU" ALT="영업조직" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnSales_org" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenConSorgCode()">&nbsp;<input NAME="txtSales_Org_nm1" TYPE="Text" MAXLENGTH="30" tag="14" size="30"></TD>
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
        <TD CLASS="TD5" NOWRAP>영업조직</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtSales_Org2" TYPE="Text" MAXLENGTH="4" tag="23XXXU" size="10" ALT="영업조직"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>영업조직명</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtSales_Org_nm2" TYPE="Text" MAXLENGTH="50" tag="22XXX" size="34" ALT="영업조직명"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>조직레벨</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtlvl" TYPE="Text" MAXLENGTH="2" tag="22XXXU" size="10" ALT="조직레벨"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>영업조직총칭</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtSales_Org_Fullnm" TYPE="Text" MAXLENGTH="70" tag="21XXX" size="70" ALT="영업조직총칭"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>영업조직영문명</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtSales_Org_Engnm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50" ALT="영업조직영문명"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>상위영업조직</TD>
        <TD CLASS="TD656">
         <input NAME="txtUpper_Sales_Org" TYPE="Text" MAXLENGTH="4" tag="21XXXU" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnSales_org" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSorgCode()">&nbsp;<input NAME="txtUpper_Sales_OrgNm" TYPE="Text" MAXLENGTH="30" tag="24" size="30"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>영업조직장명</TD>
        <TD CLASS="TD656" NOWRAP><input NAME="txtHead_usr_nm" TYPE="Text" MAXLENGTH="50" tag="21XXX" size="50" ALT="영업조직장명"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>말단조직여부</TD>
        <TD CLASS="TD656" NOWRAP>
        <input type=radio CLASS="RADIO" id=rdoEndOrgFlagY name="rdoEndOrgFlag" value="Y" tag = "21XXX" checked>
         <label for="rdoEndOrgFlagY">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
        <input type=radio CLASS = "RADIO" id=rdoEndOrgFlagN name="rdoEndOrgFlag" value="N" tag = "21XXX">
         <label for="rdoEndOrgFlagN">아니오</label></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>사용여부</TD>
        <TD CLASS="TD656" NOWRAP>
         <input type=radio CLASS="RADIO" id=rdoUsage_flag1 name="rdoUsage_flag" value="Y" tag = "21" checked>
          <label for="rdoUsage_flag1">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" id=rdoUsage_flag2 name="rdoUsage_flag" value="N" tag = "21">
          <label for="rdoUsage_flag2">아니오</label></TD>
       </TR>
       <%Call SubFillRemBodyTD656(12)%>
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
     <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck()">영업조직조회</a></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER="0" SCROLLING="NO" noresize framespacing="0" TABINDEX = -1></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>  
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtEndOrgFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX = -1>
</form>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV> 
</body>
</html>


