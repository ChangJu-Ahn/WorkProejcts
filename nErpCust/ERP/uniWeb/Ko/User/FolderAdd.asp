<%@ LANGUAGE="VBSCRIPT" %>

<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>사용자 메뉴 추가</TITLE>
<% '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################%>
<% '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* %>

<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<%'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================%>
<!-- #Include file="../inc/uni2kcm.inc" -->
<!-- #Include file="../inc/IncServer.asp" -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/EventPopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/ccm.vbs">       </Script>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incUni2KTV.vbs"></Script>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js">   </SCRIPT>



<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID          = "BizAddMenu.asp"		
Const BIZ_PGM_USERMENU_ID = "LoadUserMenu.asp"		

Const C_Sep           = "/"

Const C_IMG_Folder    = "../../CShared/image/CloseDir.gif"
Const C_IMG_Open      = "../../CShared/image/OpenDir.gif"
Const C_IMG_URL       = "../../CShared/image/Program.gif"
Const C_IMG_None      = "../../CShared/image/Program_d.gif"
Const C_IMG_Const     = "../../CShared/image/c_const.gif"

Const C_MNU_SEP       = "::"
Const C_MNU_ID        = 0
Const C_MNU_UPPER     = 1
Const C_MNU_LVL       = 2
Const C_MNU_TYPE      = 3
Const C_MNU_NM        = 4
Const C_MNU_AUTH      = 5

Const C_USER_MENU     = "사용자 메뉴"
Const C_USER_MENU_KEY = "*"
Const C_USER_MENU_STR = "UM_"
Const C_UNDERBAR      = "_"

Const C_NEW_FOLDER    = "새 폴더"

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgBlnBizLoadMenu
Dim StrVal
Dim OpMode
Dim arrParam
Dim strPgName
Dim strURL
Dim XXXX   'khy200304

arrParam  = window.dialogArguments

set XXXX = arrParam(0)'khy200304
strPgName = arrParam(1)
strURL    = arrParam(2)

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgBlnBizLoadMenu = False
'--------------------------------------------  Coding part  -----------------------------------------------
End Sub

'******************************************  2.2 화면 초기화 함수  *************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************** 
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim NodX
	Set NodX = frm2.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, C_USER_MENU, C_USFolder, C_USFolder)
	NodX.ExpandedImage = C_USOpen
	NodX.Tag = 0			
End Sub
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'   개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'==========================================================================================
'   Function Name : GetIndex
'   Function Desc : 현재 입력되는 Node가 면번째 인지를 반환한다.
'==========================================================================================
Function GetIndex(Node)
    
    If Node.Key = gDragNode.Key Then
        GetIndex = gDropNode.Children + 1
        Exit Function
    End If

	GetIndex = GetSeq(Node)    

End Function

'==========================================================================================
'   Function Name : GetNodeLvl
'   Function Desc : 현재 노드의 Level을 찾는다.
'==========================================================================================
Function GetNodeLvl(Node)
    Dim tempNode
    
    Set tempNode = Node
    
    GetNodeLvl = 0
    
    if tempNode.Key <> "*" Then
	    Do    	
    		GetNodeLvl = GetNodeLvl + 1
    		Set tempNode = tempNode.Parent
    	Loop Until tempNode.Key = "*"
	End If
	Set tempNode = Nothing
End Function

'======================================================================================================
'	메뉴를 읽어 TreeView에 넣음 
'======================================================================================================
Sub DisplayBizMenu()
	Dim strVal

	Call SetDefaultVal()
		
	frm2.uniTree1.MousePointer = 11
	
	strVal = BIZ_PGM_USERMENU_ID & "?txtKey=$"								'☆: 조회 조건 데이타 
    strVal = strVal & "&txtUKey=$"											'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(PopBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	strVal = ""

End Sub

'========================================================================================
' Function Name : GetImage
' Function Desc : 이미지 정보 
'========================================================================================
Function GetImage(Byval arrLine)
	Dim strImg
	Select Case arrLine(C_MNU_AUTH)
		Case "A"
			If arrLine(C_MNU_TYPE) = "M" Then
				strImg = C_USFolder
			Else
				strImg = C_USURL
			End If
		Case "I"
			strImg = C_USConst
		Case "N"
			strImg = C_USNone
	End Select
	GetImage = strImg
End Function

'==========================================================================================
'   Function Name : GetMenuType
'   Function Desc : 입력되는 Node의 메뉴타입을 반환한다.
'==========================================================================================
Function GetMenuType(imgStyle)
    If imgStyle = C_USURL Then
        getMenuType = "P"
    Else
        getMenuType = "M"
    End If
End Function

'==========================================================================================
'   Function Name : GetPgID
'   Function Desc : 현재 화면의 ID를 얻는다.
'==========================================================================================
Function getPgID()

'	Dim tmpPgID
	
'	tmpPgID = strURL
	
'	If InStrRev(tmpPgID, ".") > 0 Then tmpPgID = Left(tmpPgID, InStrRev(tmpPgID, ".") - 1)
'	If InStrRev(tmpPgID, "/") > 0 Then tmpPgID = Mid(tmpPgID, InStrRev(tmpPgID, "/") + 1)
	
'	getPgID = tmpPgID

	getPgID = strURL

End Function


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'==========================================================================================
'   Event Name : uniTree1_AfterLabelEdit
'   Event Desc : Add하고 Label을 수정한후 DB등록을 호출할 이벤트 
'==========================================================================================
Sub uniTree1_AfterLabelEdit(Cancel, NewString)

	Dim NodX
	
	Set NodX = frm2.uniTree1.SelectedItem
	
	StrVal = StrVal & "U"                     & gColSep 
	StrVal = StrVal & NodX.Key                & gColSep 
	StrVal = StrVal & NodX.Parent.Key         & gColSep 
	StrVal = StrVal & NewString               & gColSep 
	StrVal = StrVal & getMenuType(NodX.Image) & gColSep 
	StrVal = StrVal & GetNodeLvl(NodX)        & gColSep 
	StrVal = StrVal & NodX.Parent.Tag         & gColSep 
	StrVal = StrVal & NodX.Key                & gColSep 
	StrVal = StrVal & NodX.Parent.Key         & gRowSep
	Set NodX = Nothing	
	
	
End Sub
'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node를 클릭하면 발생 이벤트 
'==========================================================================================
Sub uniTree1_NodeClick(Node)
	if Node.Key = "*" Then
		frm2.uniTree1.LabelEdit = 1
	Else
		frm2.uniTree1.LabelEdit = 0
	End If				
End Sub

'======================================  uniTree1_onAddImgReady()  ====================================
'	Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount수의 Image가 다운로드 완료되고 TreeView의 ImageList에 추가되면 발생하는 이벤트 
'========================================================================================================= 

Sub uniTree1_onAddImgReady()
	Call DisplayBizMenu()
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	
	Dim NodX, lHwnd
	Dim IntRetCD
	
	If Trim(getPgID()) = "" Or Trim(getPgId()) = "N/A" Then 'khy200304
		IntRetCD = DisplayMsgBox("211429", "x","x","x")			    '공지 사항은 사용자 메뉴에 등록할수 없습니다.    			
    	Self.Returnvalue = ""
  		Self.Close()      			
		Exit Sub		
	End if
	
	With frm2
		.uniTree1.SetAddImageCount = 5
		.uniTree1.Indentation      = "200"	                                            ' 줄 간격 
		.uniTree1.AddImage C_IMG_Folder, C_USFolder, 0									'⊙: TreeView에 보일 이미지 지정 
		.uniTree1.AddImage C_IMG_Open  , C_USOpen  , 0
		.uniTree1.AddImage C_IMG_URL   , C_USURL   , 0
		.uniTree1.AddImage C_IMG_None  , C_USNone  , 0
		.uniTree1.AddImage C_IMG_Const , C_USConst , 0
		
		.uniTree1.OLEDragMode = 0														'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
		.uniTree1.OLEDropMode = 0

		.txtName.Value = strPgName
	End With
	
	Call InitVariables()
	
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************

'======================================================================================================
'	취소 버튼 클릭 
'======================================================================================================

Function btnCancel_onclick()
	Self.Returnvalue = ""
	Self.Close()
End Function

'======================================================================================================
'	확인 버튼 클릭시 
'======================================================================================================

Function btnOK_onclick()

	Dim NodX	
	Dim strTmp1, strTmp2, tmp_cnt
	Dim IntRetCD
	Dim NodeCnt
	
	' 폴더명을 지운다음 저장시 New folder로 들어가도록 수정 
	strTmp1 = split(StrVal,gRowSep)
	StrVal = ""
  
	For tmp_cnt = 0 to ubound(strTmp1,1) - 1
	    strTmp2 = split(strTmp1(tmp_cnt),gColSep)
	    If Trim(strTmp2(3)) > "" then
	        StrVal = StrVal & strTmp1(tmp_cnt) & gRowSep
	    End If
	Next


	If frm2.txtName.value ="" Then 		
		IntRetCD = DisplayMsgBox("211428", "x","x","x")			    '프로그램명을 입력하세요.
    			If IntRetCD = vbNo Then
      				Exit Function
    			End If	
		Exit function
	End If 

	With frm2
		Set NodX = .uniTree1.SelectedItem		 
		NodeCnt = split(NodX.fullpath,"\")               'khy
		StrVal = StrVal & "C"                  & gColSep 
		StrVal = StrVal & getPgID()            & gColSep 
		StrVal = StrVal & NodX.Key             & gColSep 
		StrVal = StrVal & .txtName.Value       & gColSep 
		StrVal = StrVal & "P"                  & gColSep  
		StrVal = StrVal & UBound(NodeCnt) + 1  & gColSep 'khy
		StrVal = StrVal & (Cint(NodX.Tag) + 1) & gColSep 
		StrVal = StrVal & ""                   & gColSep 
		StrVal = StrVal & ""                   & gRowSep

	End With


	frm2.uniTree1.MousePointer = 11
	
	'khy200304
	Dim StrRec
	Dim StrTmp
	Dim CntRec
	Dim i 
	Dim ArrRec
	
	StrTmp = Split(Strval,chr(12))
	
	CntRec = UBound(strtmp)
		
	For i = 0 To CntRec-1
		ArrRec= Split(StrTmp(i),chr(11))		
		Call XXXX.AddMenuPopup (ArrRec(0),ArrRec(1),ArrRec(2),ArrRec(3),ArrRec(4),ArrRec(5),ArrRec(6))
    	
	Next
	
	Call XXXX.RefreshUsrXml()
	
	Call DbSaveOk
	
	strval =""
	 
	Set NodX = Nothing	

End Function


'======================================================================================================
'	새폴더 버튼 클릭 
'======================================================================================================

Function btnNewFolder_onclick()

	Dim NodX, Node, tmpCnt, strNewKey
	
	Set Node = frm2.uniTree1.SelectedItem
	
	tmpCnt = 0
	
	On Error Resume Next

	'손호영 수정 2002.09.13 메뉴레벨 제한.
	'If (GetNodeLvl(Node) + 1) > 10 then
	'	IntRetCD = DisplayMsgBox("211437", "x","x","x")			    '사용자 메뉴는 10레벨을 초과할 수 없습니다.
	'	exit function
	'End If
	
	Do
		strNewKey = C_USER_MENU_STR & cStr(tmpCnt)
		Set NodX = frm2.uniTree1.Nodes(strNewKey)
        If Err.Number <> 0 Then
            Exit Do
        Else
        	Err.Clear
            tmpCnt = tmpCnt + 1
        End If
    Loop
    
    Set NodX             = frm2.uniTree1.Nodes.Add(Node.Key, tvwChild, strNewKey, C_NEW_FOLDER, C_USFolder, C_USFolder)
    NodX.ExpandedImage   = C_USOpen
    NodX.tag             = 0
	NodX.Parent.Tag      = Cint(NodX.Parent.Tag) + 1
	NodX.Parent.Expanded = True
	NodX.Selected        = True

	StrVal = StrVal &  "C"                    & gColSep
	StrVal = StrVal &  strNewKey              & gColSep
	StrVal = StrVal &  Node.Key               & gColSep
	StrVal = StrVal &  C_NEW_FOLDER           & gColSep
    StrVal = StrVal &  "M"                    & gColSep
    StrVal = StrVal &  (GetNodeLvl(Node) + 1) & gColSep
    StrVal = StrVal &  (Cint(Node.Tag) + 1)   & gColSep
    StrVal = StrVal &  ""                     & gColSep
    StrVal = StrVal &  ""                     & gRowSep
    
	Set NodX = Nothing
	Set Node = Nothing
	
	frm2.uniTree1.SetFocus
	Call frm2.uniTree1.StartLabelEdit

End Function

'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
'#########################################################################################################

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQueryOk가 성공적일때 수행 
'========================================================================================
Function DbQueryOk()
	lgBlnBizLoadMenu = True
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	frm2.uniTree1.MousePointer = 0
    Self.close
End Function
</SCRIPT>
</HEAD>

<BODY SCROLL=no rightmargin="0" bgcolor="#ffffff">
<FORM NAME="frm2" TARGET="PopBizASP" METHOD="POST">
<TABLE WIDTH="85%" HEIGHT="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
	<TR>
		<TD WIDTH="100%" VALIGN="BOTTOM">
			<TABLE WIDTH="100%" id="TABLE1">
				<TR HEIGHT=20>
					<TD WIDTH="25%"></TD>
					<TD WIDTH="75%">사용자 메뉴에 이 화면을 추가합니다.</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD VALIGN="TOP">
			<FIELDSET>
			<TABLE CELLSPACING="0" CELLPADDING="2" WIDTH="100%" id="TABLE1">
				<TR HEIGHT="25">
					<TD CLASS="TD5" WIDTH="20%" ALIGN="RIGHT"><LABEL>이름</LABEL>&nbsp;&nbsp;</TD>
					<TD CLASS="TD6" WIDTH="90%"><INPUT TYPE="TEXT" id=txtName style="font-size:9pt;HEIGHT: 100%; WIDTH: 100%" TAG =11></TD>
				</TR>
			</TABLE>
			</FIELDSET>
			<FIELDSET>
			<TABLE CELLSPACING="0" CELLPADDING="2" WIDTH="100%" id="TABLE1">
				<TR HEIGHT="20">
					<TD CLASS="TD5" WIDTH="20%" ALIGN="RIGHT" VALIGN="bottom"><LABEL>위치</LABEL>&nbsp;&nbsp;</TD>
					<TD CLASS="TD6" WIDTH="90%" ROWSPAN="18">
						<script language =javascript src='./js/folderadd_uniTree1_N237411767.js'></script>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
			</TABLE>
			</FIELDSET>
			<TABLE>
				<TR Class="Toolbar">
					<TD>
					<IFRAME ID="PopBizASP" NAME="PopBizASP" SRC='../blank.htm' width="100%" height="0" STYLE="display: ''"></IFRAME>
					</TD>
				</TR>
			</TABLE>
			<TABLE WIDTH="100%">
			 <TR>
				<TD ALIGN="CENTER">
					<IMG SRC="../image/bu_confirm_off.gif" BORDER="0" Style="CURSOR: hand" ALT="확  인" NAME="btnOK"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_confirm_on.gif',1)">&nbsp;&nbsp;
					<IMG SRC="../image/bu_cancel_off.gif"  BORDER="0" Style="CURSOR: hand" ALT="취  소" NAME="btnCancel"    onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_cancel_on.gif',1)">&nbsp;&nbsp;
					<IMG SRC="../image/bu_new_off.gif"     BORDER="0" Style="CURSOR: hand" ALT="새폴더" NAME="btnNewFolder" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_new_on.gif',1)">
				</TD>
			 </TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<TEXTAREA NAME=txtMulti STYLE="display: none"></TEXTAREA>
</FORM>
</BODY>
</HTML>
