<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : �������� ���/���� ȭ�� ó�� ASP
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************-%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim arrParent
Dim PopupParent
'Dim IsAttach
Dim szExcelData

'IsAttach = False
arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
</SCRIPT>
<%
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Dim strTitle , strMode
    Dim strTable, strStatus, intKeyNo, strSQL
    Dim strSubject, strWriter, strContents, strPasswd
    Dim arrtemp

    intKeyNo = CLng(Request("intKeyNo"))
    strMode  = CStr(Request("strMode"))							'��: Read Operation Mode (CRUD)        
%>

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<Script Language="VBScript">
'Option Explicit

Const BIZ_PGM_ID = "BDC04PB1.ASP"
Dim arFieldInfo(3)
Dim szJoinMethod
Dim nStartRow
Dim strMode
Dim arrTemp
Dim intKeyNo

strMode  = "<%= strMode %>"
arrTemp  = "<%= arrTemp %>"
intKeyNo = "<%= intKeyNo %>"

<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>
Function GetExcelText()

	Dim StrTempExcel			
	
	 '------ Check contents area ------ 
	If Not chkField(Document, "1") Then								'��: Check contents area 
		Exit Function
	End If
	Call LayerShowHide(1)
	
    StrTempExcel = ExcelBrokerControl.GetData(Trim(frm1.FileName1.value), _
                                                     CInt(nStartRow), _
                                                     arFieldInfo)
    
    If DbSave(StrTempExcel) = False Then
		Exit Function
    End If                                               
      

End function


Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function

'##########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Private Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Dim strDt
	strDt = "<%=GetSvrDate%>"
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	'frm1.tmPlanTime.text = strDt
	Call ggoOper.LockField(Document, "N")
	frm1.txtProcessID.focus
End Sub


'==========================================  3.1.2 Window_OnUnLoad() ======================================
'	Name : Window_OnUnLoad()
'	Description : Window �� �ݱ��ư(�ּ�,�ִ�ȭ��ư ���� �ִ� �ݱ��ư)�� ������ �� ����Ǵ� �κ� 
'========================================================================================================= 
Private Sub Window_OnUnLoad()
	If  window.ReturnValue <> True then
		window.ReturnValue = False
	End If
End Sub
	
'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

Function DbSave(ByVal iTempExcel)
	Dim ArrTempExcel
	
	Dim strVal
	Dim IntRows
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size
	
	Dbsave = False

	If LayerShowHide(1) = False Then Exit Function
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.PopupParent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.PopupParent.C_FORM_LIMIT_BYTE
	
	iColSep = Chr(11) : iRowSep = Chr(12)
	                                               
    ArrTempExcel =  Split(iTempExcel, iRowSep )     
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)					

	iTmpCUBufferCount = -1 
	
	strCUTotalvalLen = 0
    
    For IntRows = 0 To Ubound(ArrTempExcel) 
		strVal = ""
		strVal = ArrTempExcel(IntRows)
			    
		If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
		   Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
		   objTEXTAREA.name = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
			 
		   iTmpCUBufferMaxCount = parent.PopupParent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
		   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		   iTmpCUBufferCount = -1
		   strCUTotalvalLen  = 0
		End If
			       
		iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
		If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
		   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.PopupParent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
		   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		End If   
			         
		iTmpCUBuffer(iTmpCUBufferCount) =  strVal & iRowSep     
		strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		
    Next 
    
    If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           ' ��: Processing is OK

End Function


Function DbSaveOk()	
	Call RemovedivTextArea
	window.ReturnValue = True
	Self.Close()
End Function

'=========================================================================================================
Function OpenProcessID()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "�����ڵ�"
    arrParam(1) = "B_BDC_MASTER"
    arrParam(2) = Trim(frm1.txtProcessID.Value)
    arrParam(3) = ""
    arrParam(4) = "USE_FLAG='Y'"
    arrParam(5) = "�����ڵ�"
    
    arrField(0) = "PROCESS_ID"
    arrField(1) = "PROCESS_NAME"
    arrField(2) = "RUN_TIME"
    arrField(3) = "JOIN_METHOD"
    arrField(4) = "START_ROW"

    arrHeader(0) = "�����ڵ�"
    arrHeader(1) = "�� �� ��"
    arrHeader(2) = "����ð�"
    arrHeader(3) = "�� �� ��"
    arrHeader(4) = "����ð�"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
                                    Array(arrParam, arrField, arrHeader), _
                                    "dialogWidth=420px; dialogHeight=450px; center: Yes; " & _
                                    "help: No; resizable: No; status: No;")
    
    IsOpenPop = False

    If arrRet(0) <> "" Then
        frm1.txtProcessID.Value = Trim(arrRet(0))
        frm1.txtProcessNm.value = Trim(arrRet(1))
       ' frm1.tmPlanTime.Text = Trim(arrRet(2))
		szJoinMethod = Trim(arrRet(3))
		nStartRow = Trim(arrRet(4))
        
		Call CommonQueryRs(" FIELD_ID, SHEET_NO, FIELD_SEQ, PARENT_FIELD ", _
						   " B_BDC_FIELD ", _
						   " PROCESS_ID = '" & Trim(arrRet(0)) & "' ORDER BY SHEET_NO, FIELD_SEQ", _
						   lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		arFieldInfo(0) = lgF0
		arFieldInfo(1) = lgF1
		arFieldInfo(2) = lgF2
		arFieldInfo(3) = lgF3
    End If

    frm1.txtProcessID.focus
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

</Script>

<!-- #Include file="../../inc/uni2kcm.inc" -->
<!--
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
	<PARAM NAME="LPKPath" VALUE="../../Control/ExcelBroker.lpk">
</OBJECT>
-->

</HEAD>

<BODY BGCOLOR="#FFFFFF" SCROLL=no LEFTMARGIN=2 RIGHTMARGIN=0 TOPMARGIN=0 BOTTOMMARGIN=0>
<CENTER>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<INPUT TYPE=hidden NAME=txtMode VALUE="">
<INPUT TYPE=hidden NAME=txtMode VALUE="<%=strMode%>">
<INPUT TYPE=hidden NAME=txtKeyNo VALUE="<%=intKeyNo%>">
<INPUT TYPE=hidden name=txtFileinf VALUE="">
<INPUT TYPE=hidden name=txtFilePath VALUE="">
<TABLE CELLSPACING=0 CLASS="basicTB">

	<TR>
		<TD HEIGHT=1>&nbsp;<% ' ���� ���� %></TD>
	</TR>
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
                <TR>
                    <TD CLASS="TD5">�����ڵ�</TD>
                    <TD CLASS="TD6">
                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtProcessID" SIZE=15 MAXLENGTH=18 tag="12XXXU"  ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProcessID()">
                        <INPUT TYPE=TEXT NAME="txtProcessNm" SIZE=30 tag="14">
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">�۾���</TD>
                    <TD CLASS="TD6">
                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtJobTitle" SIZE=60 MAXLENGTH=128 tag="12"  ALT="�۾���">
                    </TD>
                </TR>
                <!--TR>
                    <TD CLASS="TD5">�����Ͻ�</TD>
                    <TD CLASS="TD6">
					    <script language =javascript src='./js/bdc04pa1_OBJECT1_tmPlanTime.js'></script>
                    </TD>
                </TR-->
<!--</FORM>-->
<!--<FORM NAME=frm2 TARGET="MyBizASP" METHOD="POST">-->
               <TR>
                    <TD CLASS="TD5">��������</TD>
                    <TD CLASS="TD6">
						<INPUT TYPE="file" NAME="FileName1" CLASS="box" SIZE="35" STYLE="ime-mode:disabled" OnKeyPress="CharNoClick()" ALT="��������" tag = "12">
                    </TD>
                </TR>
		    </TABLE>
		    </FIELDSET>
        </TD>
    </TR>
	<TR>
		<TD HEIGHT=1>&nbsp;</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="vbscript:GetExcelText()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="vbscript:FncClose()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>> 
            <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No  FRAMESPACING=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=300 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</CENTER>
<OBJECT ID="ExcelBrokerControl"
		CLASSID="CLSID:3894EE93-0291-4D97-8423-FAE813587B6E"
		CODEBASE="../../Control/ExcelBroker.CAB#version=1,1,0,64"
		WIDTH=0	HEIGHT=0>
</OBJECT>
</BODY>
</HTML>