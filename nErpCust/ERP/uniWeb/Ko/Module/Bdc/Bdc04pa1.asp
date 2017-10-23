<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 공지사항 등록/수정 화면 처리 ASP
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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
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
    strMode  = CStr(Request("strMode"))							'☜: Read Operation Mode (CRUD)        
%>

<!--
'==========================================  1.1.2 공통 Include   ======================================
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
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### %>
Function GetExcelText()

	Dim StrTempExcel			
	
	 '------ Check contents area ------ 
	If Not chkField(Document, "1") Then								'⊙: Check contents area 
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
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'**********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
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
'	Description : Window 의 닫기버튼(최소,최대화버튼 옆에 있는 닫기버튼)을 눌렀을 때 실행되는 부분 
'========================================================================================================= 
Private Sub Window_OnUnLoad()
	If  window.ReturnValue <> True then
		window.ReturnValue = False
	End If
End Sub
	
'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

Function DbSave(ByVal iTempExcel)
	Dim ArrTempExcel
	
	Dim strVal
	Dim IntRows
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
	
	Dbsave = False

	If LayerShowHide(1) = False Then Exit Function
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.PopupParent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.PopupParent.C_FORM_LIMIT_BYTE
	
	iColSep = Chr(11) : iRowSep = Chr(12)
	                                               
    ArrTempExcel =  Split(iTempExcel, iRowSep )     
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)					

	iTmpCUBufferCount = -1 
	
	strCUTotalvalLen = 0
    
    For IntRows = 0 To Ubound(ArrTempExcel) 
		strVal = ""
		strVal = ArrTempExcel(IntRows)
			    
		If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		   Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		   objTEXTAREA.name = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
			 
		   iTmpCUBufferMaxCount = parent.PopupParent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		   iTmpCUBufferCount = -1
		   strCUTotalvalLen  = 0
		End If
			       
		iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
		If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.PopupParent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		End If   
			         
		iTmpCUBuffer(iTmpCUBufferCount) =  strVal & iRowSep     
		strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		
    Next 
    
    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                                           ' ⊙: Processing is OK

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

    arrParam(0) = "업무코드"
    arrParam(1) = "B_BDC_MASTER"
    arrParam(2) = Trim(frm1.txtProcessID.Value)
    arrParam(3) = ""
    arrParam(4) = "USE_FLAG='Y'"
    arrParam(5) = "업무코드"
    
    arrField(0) = "PROCESS_ID"
    arrField(1) = "PROCESS_NAME"
    arrField(2) = "RUN_TIME"
    arrField(3) = "JOIN_METHOD"
    arrField(4) = "START_ROW"

    arrHeader(0) = "업무코드"
    arrHeader(1) = "업 무 명"
    arrHeader(2) = "실행시간"
    arrHeader(3) = "업 무 명"
    arrHeader(4) = "실행시간"
    
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
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
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
		<TD HEIGHT=1>&nbsp;<% ' 상위 여백 %></TD>
	</TR>
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
                <TR>
                    <TD CLASS="TD5">업무코드</TD>
                    <TD CLASS="TD6">
                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtProcessID" SIZE=15 MAXLENGTH=18 tag="12XXXU"  ALT="업무코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProcessID()">
                        <INPUT TYPE=TEXT NAME="txtProcessNm" SIZE=30 tag="14">
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">작업명</TD>
                    <TD CLASS="TD6">
                        <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtJobTitle" SIZE=60 MAXLENGTH=128 tag="12"  ALT="작업명">
                    </TD>
                </TR>
                <!--TR>
                    <TD CLASS="TD5">실행일시</TD>
                    <TD CLASS="TD6">
					    <script language =javascript src='./js/bdc04pa1_OBJECT1_tmPlanTime.js'></script>
                    </TD>
                </TR-->
<!--</FORM>-->
<!--<FORM NAME=frm2 TARGET="MyBizASP" METHOD="POST">-->
               <TR>
                    <TD CLASS="TD5">엑셀파일</TD>
                    <TD CLASS="TD6">
						<INPUT TYPE="file" NAME="FileName1" CLASS="box" SIZE="35" STYLE="ime-mode:disabled" OnKeyPress="CharNoClick()" ALT="엑셀파일" tag = "12">
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