<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 구매요청현황 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/22
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>구매요청현황</TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                              
Dim lgMark                                                  
Dim arrReturn					
Dim arrParam
Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
'20021227 kangjungu dynamic popup
top.document.title = PopupParent.gActivePRAspName


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3161pb1.asp"
Const C_MaxKey          = 0                                   
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------
'=============================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1
End Sub

'=============================================================================================================
Sub SetDefaultVal()
	Dim arrParam
	
	arrParam = arrParent(1)
	
	frm1.txtSoNo.value		= arrParam(0)
	frm1.txtSoSeq.value		= arrParam(1)  
	frm1.txtItem.value		= arrParam(2)  
	frm1.txtSchdNo.value	= arrParam(3) 
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=============================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("s3161qa1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock    
End Sub


'=============================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
		.vspdData.OperationMode = 5
    End With
End Sub

'=============================================================================================================
Function OKClick()
	Self.Close()
End Function
	

'=============================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=============================================================================================================
Function OpenSortPopup()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
   
   	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub


'=============================================================================================================
Function vspdData_KeyPress(KeyAscii)
   On Error Resume Next
   If KeyAscii = 27 Then
	  Call CancelClick()
   End If
End Function

'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
	If CheckRunningBizProcess = True Then Exit Sub	
			
    	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub

'=============================================================================================================
Function Document_onKeyDown()
	If window.event.keycode = 27 Then Self.Close
End Function

'=============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														

    Call DbQuery															

    FncQuery = True		
End Function

'=============================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
      	Exit Function
    End If
    
    With frm1

		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)	
		strVal = strVal & "&txtSoSeq=" & Trim(frm1.txtSoSeq.value)	
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)	
		strVal = strVal & "&txtSchdNo=" & Trim(frm1.txtSchdNo.value)	
        strVal = strVal & "&lgPageNo=" & lgPageNo                              
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True
End Function

'=============================================================================================================
Function DbQueryOk()														
	frm1.vspdData.Focus
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>	
						<TD CLASS=TD5 NOWRAP>수주번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="14XXXU" ALT="수주번호"></TD>
						<TD CLASS=TD5 NOWRAP>수주순번</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoSeq" TYPE="Text" MAXLENGTH="9" SIZE=20 tag="14XXXU" ALT="수주순번"></TD>
					</TR>	
					<TR>	
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="14XXXU" ALT="품목"></TD>
						<TD CLASS=TD5 NOWRAP>납품순번</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSchdNo" TYPE="Text" MAXLENGTH="9" SIZE=20 tag="14XXXU" ALT="납품순번"></TD>
					</TR>	
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3161pa1_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                     <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>