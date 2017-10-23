<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S3112RA7
'*  4. Program Name         : 수주내역현황(수주현황조회에서) 
'*  5. Program Desc         : 수주내역현황(수주현황조회에서) 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Kim Hyung suk
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

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

Dim lgIsOpenPop                                              

Dim lgMark                                                  
Dim lgblnWinEvent				

<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "s3112rb7.asp"     
Const C_MaxKey   = 20                 
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------	
	

'=============================================================================================================
Sub InitVariables()
	lgPageNo = ""
    lgBlnFlgChgValue = False                               'Indicates that no value changed    
    lgSortKey        = 1
End Sub

'=============================================================================================================
Sub SetDefaultVal()	
	Dim arrParam

	arrParam = arrParent(1)
	frm1.txtConSoNo.value = arrParam(0)
	frm1.txtHCur.value = arrParam(1)

	lgblnWinEvent = False
	Self.Returnvalue = ""
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'=============================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3112RA7","S","A","V20030318", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock        
End Sub

'=============================================================================================================
Sub SetSpreadLock()    
	frm1.vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.ReDraw = True
	frm1.vspdData.OperationMode = 5    
End Sub

'=============================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029																
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
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub


'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
		
    	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub

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
    
		.txtMode.value = PopupParent.UID_M0001				
		.txtHConSoNo.value = Trim(.txtConSoNo.value)		
        .lgPageNo.value = lgPageNo                     
        .lgSelectListDT.value = GetSQLSelectListDataType("A")
        .lgTailList.value = MakeSQLGroupOrderByList("A")
		.lgSelectList.value = EnCoding(GetSQLSelectList("A"))		
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
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

<BODY SCROLL=NO TABINDEX="-1">
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
						<TD CLASS=TD5>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConSoNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU" ALT="수주번호"></TD>
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
						<script language =javascript src='./js/s3112ra7_vaSpread1_vspdData.js'></script>
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
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHCur" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgPageNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgSelectListDT" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgTailList" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgSelectList" tag="14" TABINDEX="-1">

</FORM>	
		
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>