<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : H1001ma1
*  4. Program Name         : H1001ma1
*  5. Program Desc         : 기준정보관리/회사Rule등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/03
*  8. Modified date(Last)  : 2003/05/15
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : LSN
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h1028mb1.asp"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029B("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

	lgKeyStream = "1" & parent.gColSep       'You Must append one character( parent.gColSep)

End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

'    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0081", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 '   iCodeArr =  lgF0
  '  iNameArr =  lgF1
   ' Call  SetCombo2(frm1.cboFamily_type, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  AppendNumberPlace("7", "3", "0")
	Call  AppendNumberPlace("8", "2", "0")
	
	Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")
	
    Call SetDefaultVal()
	Call SetToolbar("1100100000001111")
	
	Call InitVariables
    Call InitComboBox
    call MainQuery()			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
	Call DisableToolBar( parent.TBC_QUERY)

    If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False
    Err.Clear
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    
	Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncDelete = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
	Dim strBasStrtMm
	Dim strBasStrtDd
	Dim strBasEndMm
	Dim strBasEndDd

    FncSave = False    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    Call MakeKeyStream("S")
	Call  DisableToolBar( parent.TBC_SAVE)
    If DbSave = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
            
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									     '⊙: This function lock the suitable field
    Call SetToolbar("11101000000011")
    Set	gActiveElement = document.ActiveElement   
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
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
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	With Frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
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
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    lgBlnFlgChgValue = false

	Call SetToolbar("1100100000001111")												'⊙: Set ToolBar

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
End Function

'========================================================================================================
' Name : cboFamily_type_OnChange
' Desc : developer describe this line 
'========================================================================================================
Sub txtApp_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtfirst_time_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFirst_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRest_rate_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : txtApp_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtApp_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtApp_dt.Action = 7 
        frm1.txtApp_dt.focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD BACKGROUND"../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" WIDTH="10" HEIGHT="23"></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white>주5일기준등록</font></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN="TOP">
						<TABLE <%=LR_SPACE_TYPE_60%>>
				            <TR>
				                <TD VALIGN="TOP" colspan=2></TD>
				            </TR>						
				            <TR>
				                <TD VALIGN="TOP" colspan=2>
				            		<FIELDSET CLASS="CLSFLD">
				            		<TABLE CLASS="BasicTB" CELLSPACING=0>
				            			<TR height= "10">
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>				            				
				            			</TR>				            		
				            			<TR>
				            				<TD CLASS="TD5" NOWRAP>주5일적용일</TD>
				            				<TD CLASS="TD6">
				            					<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtApp_dt name=txtApp_dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="주5일적용일" tag="21X1" VIEWASTEXT></OBJECT>');</SCRIPT>
				            				</TD>
				            				<TD CLASS="TD5" NOWRAP></TD>
				            				<TD CLASS="TD6"></TD>
				            			</TR>
				            			<TR height= "10">
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>	
				            			</TR>			            			
				            		</TABLE>
				            		</FIELDSET>
				            	</TD>
				            </TR>
                            <TR>
						        <TD VALIGN="TOP">
						            <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>주5일연장근무</LEGEND>
						            <TABLE CLASS="BasicTB" CELLSPACING=0>
				            			<TR height= "10">
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>	
				            			</TR>							            
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP>연장근무계산률</TD>
	                   						<TD CLASS="TD6">하루 8시간 이상 or 일주일 40시간이상 근무시</TD>
	                   						<TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6"></TD>
	                   					</TR>
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6">
	                   							<table>
	                   							<tr>
	                   								<td>최초</td>
	                   								<td><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFirst_time name=txtFirst_time CLASS=FPDS40 title=FPDOUBLESINGLE ALT="최초시간" tag="21X81"></OBJECT>');</SCRIPT>
	                   								<td> 시간 미만은 </td>
	                   								<td><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFirst_rate name=txtFirst_rate CLASS=FPDS65 title=FPDOUBLESINGLE ALT="최초시간적용률" tag="21X7Z"></OBJECT>');</SCRIPT> %</td>
	                   							</tr>
	                   							</table>
	                   						</TD>
	                   						<TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6"></TD>
	                   					</TR>	    
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6">이후시간은 
	                   						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRest_rate name=txtRest_rate CLASS=FPDS65 title=FPDOUBLESINGLE ALT="이후시간적용률" tag="21X7Z"></OBJECT>');</SCRIPT> % 적용</TD>
	                   						<TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6"></TD>
	                   					</TR>
				            			<TR height= "300">
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>
				            				<TD CLASS="TD5"></TD>
				            				<TD CLASS="TD6"></TD>	
				            			</TR>	
				    		        </TABLE>
				    		        </FIELDSET>
						        </TD>
                            </TR>
					    </TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

