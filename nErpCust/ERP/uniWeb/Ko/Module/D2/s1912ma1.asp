<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s1912ma1
'*  4. Program Name         : ATP설정 
'*  5. Program Desc         : ATP설정 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/05/18
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Sonbumyeol
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID = "s1912mb1.asp"												

Dim IsOpenPop						

'=================================================================================================================
Sub InitVariables()
	
	Err.Clear 
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
   
End Sub

'=================================================================================================================
Sub SetDefaultVal()
	
	Err.Clear 	
	frm1.txtconPlant_cd.focus
    frm1.rdoATP_flag1.checked = True   
	lgBlnFlgChgValue = False
End Sub

'=================================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> 
 <% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'=================================================================================================================
Function OpenConPlant(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_PLANT"                            
		arrParam(2) = Trim(frm1.txtconPlant_cd.Value)			
		arrParam(3) = ""			                            
		arrParam(4) = ""									    
		arrParam(5) = "공장"							    
	
		arrField(0) = "Plant_cd"					                
		arrField(1) = "Plant_nm"						            
    
		arrHeader(0) = "공장"						        
		arrHeader(1) = "공장명"							    

	End Select

	arrParam(0) = arrParam(5)								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSPlantDC(arrRet, iWhere)
	End If	
	
End Function

'=================================================================================================================
Function OpenCheck(strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_PLANT"                             
		arrParam(2) = Trim(frm1.txtPlant_cd.Value)			
		arrParam(3) = ""                                	
		arrParam(4) = ""									
		arrParam(5) = "공장"					     	
	
		arrField(0) = "Plant_cd"						    
		arrField(1) = "Plant_nm"							
    
		arrHeader(0) = "공장"							
		arrHeader(1) = "공장명"							

   	Case 1
		arrParam(1) = "B_MINOR"                             
		arrParam(2) = Trim(frm1.txtAtp_area_lvl.Value)	    
		arrParam(3) = ""                                	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0008", "''", "S") & ""					
		arrParam(5) = "ATP체크 레벨"					
	
		arrField(0) = "Minor_cd"							
		arrField(1) = "Minor_nm"							
    
		arrHeader(0) = "ATP체크 레벨"							
		arrHeader(1) = "ATP체크 레벨명"							



	Case 2
		arrParam(1) = "B_MINOR"                             
		arrParam(2) = Trim(frm1.txtOnhand_stk_lvl.Value)	
		arrParam(3) = ""                                	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0009", "''", "S") & ""					
		arrParam(5) = "보유재고 레벨"					
	
		arrField(0) = "Minor_cd"							
		arrField(1) = "Minor_nm"							
    
		arrHeader(0) = "보유재고 레벨"					
		arrHeader(1) = "보유재고 레벨명"				

	
	Case 3
		arrParam(1) = "B_MINOR"                             
		arrParam(2) = Trim(frm1.txtPlaned_gi_lvl.Value)		
		arrParam(3) = ""                                	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0010", "''", "S") & ""					
		arrParam(5) = "출고예정 레벨"				    
	
		arrField(0) = "Minor_cd"							
		arrField(1) = "Minor_nm"							
    
		arrHeader(0) = "출고예정 레벨"					
		arrHeader(1) = "출고예정 레벨명"				

	Case 4
		If frm1.txtAtp_area_lvl.value = "" Then 
	        Call DisplayMsgBox("200914","x","x","x")
			IsOpenPop = False
			Exit Function
		End If			 
		
		arrParam(1) = "b_configuration  C, b_minor B"       
		arrParam(2) = Trim(frm1.txtPlaned_gr_lvl.Value)		
		arrParam(3) = ""                                	
		arrParam(4) = "C.major_cd = B.major_cd and C.minor_cd = B.minor_cd and B.major_cd = " & FilterVar("S0011", "''", "S") & " and C.reference =  " & FilterVar(frm1.txtAtp_area_lvl.Value, "''", "S") & ""					
		arrParam(5) = "입고예정 레벨"					
	
		arrField(0) = "B.minor_cd"							
		arrField(1) = "B.minor_nm"							
    
		arrHeader(0) = "입고예정 레벨"				    
		arrHeader(1) = "입고예정 레벨명"				
	
	End Select

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOpenCheck(arrRet, iWhere)
	End If	
	
End Function

'=============================================================================================================
Function SetConSPlantDC(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		Case 0
			.txtconPlant_cd.value = arrRet(0) 
			.txtconPlant_nm.value = arrRet(1) 
			.txtconPlant_cd.focus  
		End Select
	End With
	
End Function

'=============================================================================================================
Function SetOpenCheck(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		Case 0
			.txtPlant_cd.value = arrRet(0) 
			.txtPlant_nm.value = arrRet(1)   
			.txtPlant_cd.focus
		Case 1
			.txtAtp_area_lvl.value = arrRet(0) 
			.txtAtp_area_lvl_nm.value = arrRet(1) 
			.txtAtp_area_lvl.focus  
			Call txtAtp_area_lvl_OnChange()
		Case 2
			.txtOnhand_stk_lvl.value = arrRet(0) 
			.txtOnhand_stk_lvl_nm.value = arrRet(1)   
			.txtOnhand_stk_lvl.focus  
		Case 3
			.txtPlaned_gi_lvl.value = arrRet(0) 
			.txtPlaned_gi_lvl_nm.value = arrRet(1)   
			.txtPlaned_gi_lvl.focus  
		Case 4
			.txtPlaned_gr_lvl.value = arrRet(0) 
			.txtPlaned_gr_lvl_nm.value = arrRet(1)   
			.txtPlaned_gr_lvl.focus  
		End Select

		lgBlnFlgChgValue = True

	End With
	
End Function

'==============================================================================================================
Sub CookiePage(Byval Kubun)

	On Error Resume Next

	Err.Clear 
	
	Const CookieSplit = 4877						

	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtconPlant_cd.value 
		
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
		
		If strTemp = "" then Exit Sub	

		arrVal = Split(strTemp, gRowSep)
		
		If arrVal(0) = "" Then Exit Sub

		frm1.txtconPlant_cd.value =  arrVal(0)
	
		If Err.number <> 0 Then
			Err.Clear 
			WriteCookie CookieSplit , ""
			Exit Sub
		End If
		
		Call MainQuery()
			
		WriteCookie CookieSplit , ""

	End IF
	
End Sub

'==============================================================================================================
 Sub Form_Load()

	Call LoadInfTB19029															
    Call ggoOper.LockField(Document, "N")                                   
	Call SetDefaultVal
	Call InitVariables														

    Call SetToolbar("1110100000001111")										
	Call CookiePage(0)

End Sub

'==============================================================================================================
Sub rdoATP_flag1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoATP_flag2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub
'==============================================================================================================
Sub txtAtp_area_lvl_OnChange()
	frm1.txtPlaned_gr_lvl.value = ""
	frm1.txtPlaned_gr_lvl_nm.value = ""	
End Sub

'==============================================================================================================
 Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
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

    Call DbQuery																

    FncQuery = True																
        
End Function

'===============================================================================================================
 Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "A")                                      	  						
    Call ggoOper.LockField(Document, "N")                                       
    Call SetDefaultVal
    Call InitVariables															
    Call SetToolbar("11101000000011")

    FncNew = True																

End Function

'===============================================================================================================
 Function FncDelete() 
    
    Dim IntRetCD
    
    FncDelete = False														
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")        
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO,"x","x")

    If IntRetCD = vbNo then exit function
    
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
        
    If frm1.rdoATP_flag1.checked =true then
       frm1.txtRadioflag.value=frm1.rdoATP_flag1.value
    ElseIF frm1.rdoATP_flag2.checked =true then
       frm1.txtRadioflag.value=frm1.rdoATP_flag2.value
    End If 
        
    CAll DbSave				                                                
    
    FncSave = True                                                          
    
End Function

'===============================================================================================================
 Function FncCopy() 
	Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE			
        
    Call ggoOper.ClearField(Document, "1")                                      
    Call ggoOper.LockField(Document, "N")									

    Call InitVariables	
    Call SetToolbar("11101000000111")
    
   
    frm1.txtPlant_cd.value = "" 
    frm1.txtPlant_nm.value = "" 
    frm1.txtPlant_cd.focus
    
End Function

'===============================================================================================================
 Function FncCancel() 
    On Error Resume Next                                                    
End Function

'===============================================================================================================
 Function FncInsertRow() 
     On Error Resume Next                                                   
End Function

'===============================================================================================================
 Function FncDeleteRow() 
    On Error Resume Next                                                    
End Function

'===============================================================================================================
 Function FncPrint() 
    Call parent.FncPrint()
End Function

'===============================================================================================================
 Function FncPrev() 
    On Error Resume Next                                                    
End Function

'===============================================================================================================
 Function FncNext() 
    On Error Resume Next                                                    
End Function

'===============================================================================================================
 Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)
End Function

'===============================================================================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   '☜:화면 유형, Tab 유무 
End Function

'===============================================================================================================
 Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD =DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    FncExit = True
End Function

'===============================================================================================================
 Function DbDelete() 
    Err.Clear																		
    
    DbDelete = False																

	
	If   LayerShowHide(1) = False Then
	     Exit Function 
	End If
	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003									
    strVal = strVal & "&txtconPlant_cd=" & Trim(frm1.txtconPlant_cd.value)			

	Call RunMyBizASP(MyBizASP, strVal)												
	
    DbDelete = True																	

End Function

'===============================================================================================================
Function DbDeleteOk()														
	Call FncNew()
End Function

'===============================================================================================================
 Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         
	
	
	If   LayerShowHide(1) = False Then
	     Exit Function 
	End If


	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
    strVal = strVal & "&txtconPlant_cd=" & Trim(frm1.txtconPlant_cd.value)			
   
	Call RunMyBizASP(MyBizASP, strVal)												
	
    DbQuery = True																	

End Function
'===============================================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									
	Call SetToolbar("1111100000111111")

End Function

'===============================================================================================================
 Function DbSave() 

    Err.Clear																

	DbSave = False															

	
	If   LayerShowHide(1) = False Then
	     Exit Function 
	End If



	With frm1
		.txtMode.value = parent.UID_M0002											
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID 
		.txtUpdtUserId.value = parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
	
    DbSave = True                                                           
    
End Function

'===============================================================================================================
Function DbSaveOk()															

	With frm1
		.txtconPlant_cd.value = .txtPlant_cd.value 
		
	End With
	
    Call InitVariables
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ATP설정</font></td>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconPlant_cd" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPlant frm1.txtconPlant_cd.value,0">&nbsp;
									                <INPUT NAME="txtconPlant_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>	
			                        <TD CLASS="TDT"></TD>
								    <TD CLASS="TD6"></TD>					
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
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD656><INPUT NAME="txtPlant_cd" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCheck frm1.txtPlant_cd.value,0">&nbsp;
								                <INPUT NAME="txtPlant_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="24"></TD>
							</TR>
						    <TR>
							     <TD CLASS=TD5 NOWRAP>ATP사용</TD>
								 <TD CLASS=TD656 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoATP_flag" id="rdoATP_flag1" value="Y" tag = "21XXX" checked>
										<label for="rdoATP_flag1">예</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoATP_flag" id="rdoATP_flag2" value="N" tag = "21XXX">
										<label for="rdoATP_flag2">아니오</label>
								 </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ATP 체크레벨</TD>
								<TD CLASS=TD656><INPUT NAME="txtAtp_area_lvl" ALT="ATP체크레벨" TYPE="Text" MAXLENGTH="5" SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCheck frm1.txtAtp_area_lvl.value,1">&nbsp;
								                <INPUT NAME="txtAtp_area_lvl_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
							</TR>		
							<TR>
								<TD CLASS=TD5 NOWRAP>보유재고레벨</TD>
								<TD CLASS=TD656><INPUT NAME="txtOnhand_stk_lvl" ALT="보유재고레벨" TYPE="Text" MAXLENGTH="5" SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCheck frm1.txtOnhand_stk_lvl.value,2">&nbsp;
								                <INPUT NAME="txtOnhand_stk_lvl_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>출고예정레벨</TD>
								<TD CLASS=TD656><INPUT NAME="txtPlaned_gi_lvl" ALT="출고예정레벨" TYPE="Text" MAXLENGTH="5" SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCheck frm1.txtPlaned_gi_lvl.value,3">&nbsp;
								                <INPUT NAME="txtPlaned_gi_lvl_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>입고예정레벨</TD>
								<TD CLASS=TD656><INPUT NAME="txtPlaned_gr_lvl" ALT="입고예정레벨" TYPE="Text" MAXLENGTH="5" SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCheck frm1.txtPlaned_gr_lvl.value,4">&nbsp;
								                <INPUT NAME="txtPlaned_gr_lvl_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
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
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioflag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
