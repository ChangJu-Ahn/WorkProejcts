<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1212MA2
'*  4. Program Name         : ����óĮ���ټ��� 
'*  5. Program Desc         : ����óĮ���ټ��� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/01/16
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
' ���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<STYLE TYPE="text/css">
 .Header {height:24; font-weight:bold; text-align:center; color:darkblue}
 .Day {height:22;cursor:Hand;
  font-size:17; font-weight:bold; Border:0; text-align:right}
 .DummyDay {height:22;cursor:;
  font-size:12; font-weight:; Border:0; text-align:right}
</STYLE>
<MAP NAME="CalButton">
 <AREA SHAPE=RECT COORDS="1, 1, 20, 20" ALT="Year -" onClick="ChangeMonth(-12)">
 <AREA SHAPE=RECT COORDS="20, 1, 40, 20" ALT="Month -" onClick="ChangeMonth(-1)">
 <AREA SHAPE=RECT COORDS="40, 1, 60, 20" ALT="Month +" onClick="ChangeMonth(1)">
 <AREA SHAPE=RECT COORDS="60, 1, 80, 20" ALT="Year +" onClick="ChangeMonth(12)">
</MAP>

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                      
<!-- #Include file="../../inc/lgvariables.inc" -->
<!--'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================!-->

Const BIZ_PGM_ID = "M1212mb2.asp"
Const CChnageColor = "#f0fff0"
<!-- '==========================================  1.2.2 Global ���� ����  =====================================
' 1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
' 2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= !-->
Dim lgIsOpenPop     
Dim lgNextNo     
Dim lgPrevNo     
<!-- '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= !-->
<!-- '----------------  ���� Global ������ ����  ----------------------------------------------------------- !-->

<!-- '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ !-->
Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)

<!-- '==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= !-->
Sub InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE   
	lgBlnFlgChgValue = False    
	lgIntGrpCount = 0           

	Dim iRow, iCol
	For iRow = 1 To 6
		For iCol = 1 To 7
			If frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = CChnageColor Then
				frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = "white"
				frm1.txtDate((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtDesc((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
			End If
		Next
	Next
End Sub

<!--'=============================== 2.1.2 LoadInfTB19029() =================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== !-->
 Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

<!-- '==========================================  2.2.1 InitComboBox()  ========================================
' Name : InitComboBox()
' Description : ComboBox �ʱ�ȭ 
'========================================================================================================= !-->
 Sub InitComboBox()
	Dim i, ii
	Dim oOption
 
	For i=<%=Year(GetSvrDate)%>-10 To <%=Year(GetSvrDate)%>+10
		Call SetCombo(frm1.cboYear, i, i)
	Next

	If Len(ReadCookie ("Year")) Then
		frm1.cboYear.value = ReadCookie ("Year")
		WriteCookie "Year",""
	Else
		frm1.cboYear.value = <%=Year(GetSvrDate)%>
	End If
		 
	For i=1 To 12
		ii = Right("0" & i, 2)
		Call SetCombo(frm1.cboMonth, ii, ii)
	Next

	frm1.cboMonth.value = Right("0" & "<%=Month(GetSvrDate)%>", 2)
End Sub
<!-- '==========================================  2.2.2 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= !-->
Sub SetDefaultVal()
	frm1.txtBpCd.focus 
	Set gActiveElement = document.activeElement
	Call SetToolbar("1100100000001111")    
End Sub

<!-- '------------------------------------------  OpenBpCd()  -------------------------------------------------
' Name : OpenBpCd()
' Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����ó"    
	arrParam(1) = "B_Biz_Partner"   
	arrParam(2) = Trim(frm1.txtBpCd.Value) 
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " " 
	arrParam(5) = "����ó"      
	 
	arrField(0) = "BP_CD"       
	arrField(1) = "BP_NM"       
	    
	arrHeader(0) = "����ó"      
	arrHeader(1) = "����ó��"     
	    
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

<!-- '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= !-->
Sub Form_Load()
	Call InitComboBox      
	Call LoadInfTB19029        
	Call ggoOper.LockField(Document, "N")           
	Call InitVariables        
	Call SetToolbar("1100100000001111")    
	Call SetDefaultVal
	  
	If ReadCookie("Bp_Cd")<>"" then     
		frm1.txtBpCd.value = ReadCookie("Bp_Cd")
		Call MainQuery()
		parent.WriteCookie "Bp_Cd",""
	End If
End Sub

Sub DescChange(iDate)
	Dim strDesc
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If
	 
	Call SetChange(iDate)

	strDesc = frm1.txtDesc(index).value
	frm1.txtDesc(index).value = ""
	 
	frm1.txtDesc(index).value = strDesc
	frm1.txtDesc(index).title = strDesc
End Sub

Sub HoliChange(iDate)
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If

	Call SetChange(iDate)
	 
	If frm1.txtHoli(index).value = "H" Then
		If (index+1) Mod 7 = 0 Then
			frm1.txtDate(index).style.color = "blue"
		Else
			frm1.txtDate(index).style.color = "black"
		End If
		frm1.txtHoli(index).value = "D"
	Else
		frm1.txtDate(index).style.color = "red"
		frm1.txtHoli(index).value = "H"
	End if
End Sub

Sub SetChange(iDate)
	Dim index
	index = iDate - 1

	lgBlnFlgChgValue = True
	 
	frm1.All.tblCal.Rows(Int((index+7)/7)).Cells(index Mod 7).Style.backgroundColor = CChnageColor
	frm1.txtDate(index).Style.backgroundColor = CChnageColor
	frm1.txtDesc(index).Style.backgroundColor = CChnageColor
End Sub

Sub ChangeMonth(i)
    Dim strVal
    Dim dtDate
    Dim IntRetCD

    '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
	End If

	Call InitVariables 
	 
	On Error Resume Next
	Err.Clear
	 
	dtDate = CDate(frm1.hYear.value & "-" & frm1.hMonth.value & "-" & "01")

	If Err.Number <> 0 Then                         
		Err.Clear
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Sub
	End If

	dtDate = DateAdd("m", i, dtDate)
	 
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       '��: 
	strVal = strVal & "&txtYear=" & Year(dtDate)       '��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtMonth=" & Month(dtDate)       '��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.Value)       '��: ��ȸ ���� ����Ÿ 
	 
	If LayerShowHide(1) = False then
		Exit Sub
	end if
	Call RunMyBizASP(MyBizASP, strVal)
End Sub

<!--
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
-->
Function FncQuery() 
	Dim IntRetCD 

	if Not chkField(Document, "1") Then    
		Exit Function
	End If
	
	FncQuery = False                                
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	'-----------------------
	'Erase contents area
	'-----------------------
	Call InitVariables        
	    
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then Exit Function
	       
	FncQuery = True         
End Function

<!--
'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
-->
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                 
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
       
    End If
    
    If DbSave = False Then Exit Function
    
    FncSave = True                                  
End Function

<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
 Function FncPrint() 
    On Error Resume Next              
    Call parent.FncPrint()
    
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
 Function FncExit()
	Dim IntRetCD
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    FncExit = True
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
-->
 Function DbQuery() 
	Dim strVal
    DbQuery = False                                    
	    
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001    
	strVal = strVal & "&txtYear=" & Trim(frm1.cboYear.value) 
	strVal = strVal & "&txtMonth=" & Trim(frm1.cboMonth.Value) 
	strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.Value) 
	   
	if LayerShowHide(1) = False then
		Exit Function 
	end if
	Call RunMyBizASP(MyBizASP, strVal)       
	DbQuery = True                                              
End Function

<!--
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
-->
Function DbQueryOk()           
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE         
End Function

<!--
'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
-->
Function DbSave() 
	Err.Clear          

	DbSave = False         

	frm1.txtMode.value = parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	 
	if LayerShowHide(1) = False then
		Exit Function 
	end if
	    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)    

	DbSave = True                                   
End Function

<!--
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
-->
Function DbSaveOk()         
    Call InitVariables()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB">
 <TR>
  <TD HEIGHT=5>&nbsp;<% ' ���� ���� %></TD>
 </TR>
 <TR HEIGHT=23>
  <TD WIDTH=100%>
   <TABLE CLASS="BasicTB" CELLSPACING=0>
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����óĮ���ټ���</font></td>
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
   <TABLE WIDTH=99% HEIGHT=100% BORDER=0 CELLSPACING=0 ALIGN="CENTER">
    <TR>
     <TD HEIGHT=5 WIDTH=100%></TD>
    </TR>
    <TR>
     <TD>
      <TABLE ID="tbTitle" WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="center">
       <TR>
         <TD CLASS="TD5" NOWRAP>����ó</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
                 <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14X"></TD>   
         <TD CLASS="TD5">�������</TD>
         
        <TD WIDTH=10>
         <SELECT Name="cboYear" tag="22" CLASS=cboNormal></SELECT>
        </TD>
        <TD CLASS="TD6">
         <SELECT Name="cboMonth" tag="22" CLASS=cboNormal></SELECT>
        </TD>
        <TD CLASS="TD6" NOWRAP></TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
    <TR>
     <TD>
      <TABLE ID="tblCal" WIDTH=100% HEIGHT=100% BORDER=1 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
       <THEAD CLASS="Header">
        <TR>
         <TD>�Ͽ���</TD>
         <TD>������</TD>
         <TD>ȭ����</TD>
         <TD>������</TD>
         <TD>�����</TD>
         <TD>�ݿ���</TD>
         <TD>�����</TD>
                 </TR>
             </THEAD>
       <TBODY>
<%
Dim i, j, k
k = 1
For i=1 To 6
%>
                 <TR>
<%
 For j=1 To 7
%>
         <TD ALIGN="Center">
          <TABLE WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="Center">
           <TR>
            <TD ALIGN="Left">
             <INPUT type="hidden" name="txtHoli" size=1 maxlength=1 disabled tag=2>
             <INPUT type="text" name="txtDate" class="DummyDay" size=2 maxlength=2  
              tabindex=-1 readonly disabled tag=2 onclick="HoliChange(<%=k%>)">
            </TD>
           </TR>
          </TABLE>
          <INPUT type="text" name="txtDesc" Maxlength=7 size=7 Style="Width:100;Border:0;text-align:center" disabled tag=2 onchange="DescChange(<%=k%>)">
         </TD>
<%
  k = k + 1
 Next
%>
        </TR>
<%
Next
%>
       </TBODY>
      </TABLE>
     </TD>
    </TR>
    <TR>
     <TD HEIGHT=5 WIDTH=100%></TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1501mb1.asp" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hYear" tag="24">
<INPUT TYPE=HIDDEN NAME="hMonth" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
