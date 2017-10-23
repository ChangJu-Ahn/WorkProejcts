<% Option Explicit %>
<!--
'**********************************************************************************************
'*  1. Module Name          : SALE
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit																	'☜: indicates that All variables must be declared in advance
Dim arrParent
Dim popupParent
Dim strMode
Dim strTable, strStatus, project_code, strSQL
Dim strReport_No, strins_person, strReport_Text, strPasswd
Dim arrtemp

arrParent   = window.dialogArguments
Set popupParent = arrParent(0)

</SCRIPT>

<%

Dim project_code,strMode,lngRow

    '---------------------------------------Common-----------------------------------------------------------
    Call LoadBasisGlobalInf()  
    Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
  
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim strSystemFolder
	Dim iTempPath
	Dim strReport_No,strins_person,strReport_Text
	Dim arrTemp
	Dim strReport_Abbr
	Dim struse_dt
	Dim strIns_dt
	Dim strReport_Nobar
	
	project_code = CStr(UCASE(Request("project_code")))
	strReport_No = CStr(UCASE(Request("Report_no")))
	
	strMode  = CStr(Request("strMode"))	

	strSystemFolder = GetSpecialFolder(0) '0->WindowsFolder, 1->SystemFolder, 2->TemporaryFolder		
	strSystemFolder = strSystemFolder & "\TEMP"
	
	If right(strSystemFolder,1) <> "\" Then
		iTempPath = strSystemFolder & "\UNIERPTEMP\"
	Else
		iTempPath = strSystemFolder & "UNIERPTEMP\"
	End If

	'TEMP폴더 없으면 생성 
	Call CreateFolder(iTempPath)
	Response.Cookies("unierp")("gTempDirForFileUpload") = Replace(iTempPath, "\", "/")
	
	Select Case strMode
		Case CStr(UID_M0001) 
			strReport_Nobar="등록"   
	    Case CStr(UID_M0002) 
			Call  SubBizQueryMulti()	
			strReport_Nobar="수정"                                             
	End Select

'-----------------------------------------------------------------------------------------
Sub SubBizQueryMulti()
'-----------------------------------------------------------------------------------------
    Dim strData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim arr,arrCnt
    Dim i,j,StrSQL,kk,adoRec
    'BlankchkFlg = False
    
    On Error Resume Next  
	Err.Clear                                                   '☜: Clear Error status
 	
	'project_code="2"
	LngRow = 0
	StrSQL = "SELECT ins_user,report_no,report_abbr,report_text,ins_dt"
	StrSQL = StrSQL & " FROM S_PRJ_REPORT_HDR_KO412 (NoLock) WHERE project_code = '" & project_code & "' and report_no = '" & strReport_No & "'"

	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)    
	         
		If 	FncOpenRs("R",lgObjConn,lgObjRs,StrSQL,"X","X") = False Then                    'If data not exists	

				'Call DisplayMsgBox("210100", vbInformation, "", "", I_MKSCRIPT)      '☜ : 사용자 정보 관리에 해당하는 자료가 존재하지 않습니다.
				'Call SubCloseDB(lgObjConn)
				'Response.End   
		Else
				arrCnt = lgObjRs.RecordCount 
				arr=lgObjRs.GetRows
				Call SubCloseDB(lgObjConn)  
		
				strins_person	= ConvSPChars(arr(0,0))
				strReport_No	= ConvSPChars(arr(1,0))
				strReport_Abbr	= ConvSPChars(arr(2,0))
				strReport_Text	= arr(3,0)	'arr(3,0)
				strIns_dt		= UniConvDate(arr(4,0))
		
	    End If

End Sub    

Function GetSpecialFolder(iDx)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   GetSpecialFolder = pfile.GetSpecialFolder(CInt(iDx))   
   Set pfile = Nothing
End Function
		
Function CreateFolder(iTempPath)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   Call pfile.CreateFolder(iTempPath)   
   Set pfile = Nothing
End Function


%>
<html>
<head>
<title><%=strReport_NoBar%></title>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->


<Script Language="VBScript">
'Option Explicit

Const BIZ_PGM_ID = "S3322PB1_KO412.asp"

Dim arrFileinf
Dim pStrFileInfo
Dim strMode
Dim arrTemp
Dim project_code
strMode  = "<%=strMode %>"
arrTemp  = "<%=arrTemp %>"
project_code = "<%=project_code %>"

<%'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
%>
<%
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
%>

	
<% '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
%>
function CheckValid()
	'-----------------------
    'Check condition area
    '-----------------------
  
   Call LayerShowHide(1)

'   if UNIConvDateCompanyToDB(frm1.use_dt.text,"") < UNIConvDateCompanyToDB( frm1.insrt_dt.value,"") then
'		Call DisplayMsgBox("971012", "X", "유효일", "X")
'		frm1.use_dt.focus()
'		exit function
'  end if
   
    If Not chkField(Document, "A") Then         
       Exit Function
    End If

    call dbSave()
End function

Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 

Private Sub Form_Load()

    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, popupParent.gDateFormat, popupParent.gComNum1000, popupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N") 
	call SetDefaultVal()

	If CStr(strMode) = "<%=CStr(UID_M0002)%>" Then	'글 등록이 아니면, 즉 글수정이면,
		Call FncQueryFileInfo()
		'frm1.use_dt.text  = UniConvDateAToB("<%=strUse_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=strIns_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	else
	
		'frm1.use_dt.text  = UniConvDateAToB("<%=dateadd("m",3,GetSvrDate)%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=GetSvrDate%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	End If

	Call InitComboBox

	frm1.txtTitle.value="<%=strReport_No%>"
	frm1.report_abbr.value="<%=strReport_Abbr%>"
	frm1.ins_person.value="<%=strins_person%>"
	Call ggoOper.SetReqAttr(frm1.insrt_dt, "Q")
	
End Sub

'==========================================  2.2.6 InitComboBox()  ========================================
' Name : InitComboBox()
' Desc : Combo Display
'==========================================================================================================
Sub InitComboBox()

	'// 품의제목
	Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD = " & FilterVar("SX006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.txtTitle , lgF0, lgF1, Chr(11))
	
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.ins_person.value=popupParent.gUsrID
End Sub

Function FncQueryFileInfo()

	''On Error Resume Next
	Dim IntRetCD1
	Dim iLngRow
	Dim strData
	Dim strRet
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim ArrTmpF0,ArrTmpF1,ArrTmpF2,ArrTmpF3,ArrTmpF4,ArrTmpF5,ArrTmpF6
	
	IntRetCD1= CommonQueryRs("REPORT_NM,REPORT_ID, REPORT_SIZE ","S_PRJ_REPORT_DTL_KO412", "project_code = '" & project_code & "' and report_no = '<%=strReport_No%>'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	ArrTmpF0 = split(lgF0,popupParent.gColSep)	
	ArrTmpF1 = split(lgF1,popupParent.gColSep)	
	ArrTmpF2 = split(lgF2,popupParent.gColSep)	

	strData = ""
	strRet = ""
	For iLngRow = 0 To UBound(ArrTmpF0, 1) - 1
		strData = "1" & ArrTmpF0(iLngRow) & "" & ArrTmpF2(iLngRow) & "" & ArrTmpF1(iLngRow) & "101344919970601092656N0NNFFY00YI"		
		strRet = strRet & mid(strData,3,len(strData))
		Call SetAttachFile(strData)
	Next
	Call MakeFileInfoArray("1" & strRet)		
End Function


'==========================================  3.1.2 Window_OnUnLoad() ======================================
'	Name : Window_OnUnLoad()
'	Description : Window 의 닫기버튼(최소,최대화버튼 옆에 있는 닫기버튼)을 눌렀을 때 실행되는 부분 
'========================================================================================================= 
Private Sub Window_OnUnLoad()
	If  window.ReturnValue <> True then
		window.ReturnValue = False
	End If
End Sub
	

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
Function DbSave()

	If UCASE(frm1.txtblnFileAttached.value) = "TRUE" Then
		Call MakeFileInfoString(arrFileinf)
	End If
	Call LayerShowHide(1) 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	

End Function

Function DbSaveOk()	

	//parent.frm1.vspdData.MaxRows = 0
	window.ReturnValue = True
	Self.Close()
End Function


Function vbAttachFile()

	Dim strRet

	strRet = AttachFile()

	if len(strRet) > 3  then

		Call MakeFileInfoArray(strRet)

	else
	exit function
		
	end if
	

End Function

Function MakeFileInfoArray(strRet)
	Dim arrTemp,arrTemp2
	Dim i,j
	Dim iFileCount
	Dim iCurrentSize
	
	arrTemp = Split(strRet,Chr(31))
	
	iFileCount = UBound(arrTemp) - 1	'iFileCount	첨부되는 파일 총갯수	
	
	If iFileCount = 0 Then
		Exit Function
	End If
	frm1.txtblnFileAttached.value = "TRUE"

	If IsArray(arrFileinf) Then
		iCurrentSize = UBound(arrFileinf,2)
		Redim Preserve arrFileinf(21, Cint(iCurrentSize + iFileCount))			
	Else
		iCurrentSize = 0
		Redim arrFileinf(21, iCurrentSize + iFileCount)
	End If
	For i = 1 To iFileCount				'arrTemp(0)은 첨부되는 파일 총갯수 + 1
		
		arrTemp2 = Split(arrTemp(i),Chr(29))
		
		For j = 1 To UBound(arrTemp2) + 1		
			arrFileinf(j, iCurrentSize + i) = arrTemp2(j - 1)
		Next
	Next	
	
	Call ShowArrayString(arrFileinf)

End Function

Function MakeFileInfoString(prArrTemp)
	Dim i

	Const C_FileName = 2
	Const C_FileId = 4	
	Const C_FileSize = 9
	Const C_FileCDate = 10
	
	For i = 1 To UBound(prArrTemp,2)	
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileName, i)  & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileId, i)    & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileSize, i)  & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileCDate, i) & popupParent.gRowSep
	Next
	frm1.txtFileinf.value = pStrFileInfo
	
End Function

dim deltemp
deltemp=""

Function vbDeleteFile
	Dim iDx
	Dim i
	Dim j
	Dim temp
	
	If frm1.filelist.length > 0 Then
	
		If frm1.filelist.selectedIndex < 0 Or IsNull(frm1.filelist.selectedIndex) Then 
			MsgBox "삭제할 파일을 선택하십시오."
			Exit Function
		End If
		
		For iDx = CInt(frm1.filelist.length) - 1 To 0 Step -1
			If frm1.filelist.options(iDx).Selected Then
				
				temp=frm1.filelist.options(iDx).value
				if len(deltemp)>1 then
					deltemp=deltemp &","& split(temp,"")(3)
				else 
					deltemp= split(temp,"")(3)
				end if
				
				
				frm1.delTemp.value =	deltemp
	
				frm1.filelist.remove(iDx)
				RearrangeArray(iDx + 1)	
			End If
		Next	

	Else		
		Exit Function
	End If
End Function

Sub RearrangeArray(iDx)
	' 배열을 재정돈하고 맨뒤의 배열은 dispose한다.
    Dim i,j

    For i = iDx To UBound(arrFileinf,2) - 1
        For j = 1 To UBound(arrFileinf,1)	'21개 고정 
			arrFileinf(j, i) = arrFileinf(j, i + 1)
        Next        
    Next

	Redim Preserve arrFileinf(UBound(arrFileinf,1), UBound(arrFileinf,2) - 1)

End Sub

Function ShowArrayString(prArrTemp)
	Dim i,j,iStrTemp
	
	For i = 1 To UBound(prArrTemp,2)	
		For j = 1 To UBound(prArrTemp,1)
			iStrTemp =  iStrTemp & "(" & j & "," & i & ") = " & prArrTemp(j,i) & vbCrlf
		Next
	Next
	
	
End Function

Function FetchWebSvrIp()	

	Dim gHttpWebSvrIPURL
	
	gHttpWebSvrIPURL =  "http://<%= request.servervariables("server_name") %>"	
	FetchWebSvrIp = Split(gHttpWebSvrIPURL, "/")(2)
End Function


</Script>

<script language=javascript>
function AttachFile(){
	var sRet;
	var optStr;
	var newOpt;
	var temp;
	var temp2;
	var strWebSvrIp;

	document.FR_ATTWIZ.SetLanguage('');
	document.FR_ATTWIZ.SetServerOption(0,0);
	document.FR_ATTWIZ.SetUploadMode(1);
	document.FR_ATTWIZ.SetFileUIMode(1);
	document.FR_ATTWIZ.SetModUpload();
	document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	document.FR_ATTWIZ.SetTempDir("<%=Request.Cookies("unierp")("gTempDirForFileUpload")%>");
	sRet = document.FR_ATTWIZ.AttachFile('NEW');

	var arrTemp = sRet.split(String.fromCharCode(31));

	for(var i = 0; i < arrTemp.length; i++){

		var iFileinf = arrTemp[i].split(String.fromCharCode(29));

		if(iFileinf.length > 2){
			optStr = new Option(iFileinf[1],sRet,true);	
			document.frm1.filelist.options[document.frm1.filelist.length] = optStr;
		}
	}
	//document.frm1.FileInfo.value  = sRet; 
	return(sRet);
}

function SetAttachFile(sRet){
	var arrTemp = sRet.split(String.fromCharCode(31));
	
	for(var i = 0; i < arrTemp.length; i++){		
		var iFileinf = arrTemp[i].split(String.fromCharCode(29));
		
		if(iFileinf.length > 2){
			optStr = new Option(iFileinf[1],sRet,true);				
			document.frm1.filelist.options[document.frm1.filelist.length] = optStr;
		}
	}		
}
</script>

<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</head>
<BODY TABINDEX="-1" SCROLL="no">


<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<input type=hidden name=txtMode value="<%=strMode%>">
<input type=hidden name=txtblnFileAttached value="">
<input type=hidden name=txtFileinf value="">
<input type=hidden name=project_code value="<%=project_code%>">

<input type=hidden name=delTemp value="">
<TABLE  CELLSPACING=0 CELLPADDING=0 >
	
	
	<TR HEIGHT=100%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD  HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
					<TABLE  CLASS="BasicTB" CELLSPACING=0>
						<TR>
							<TD CLASS="TD5" NOWRAP>프로젝트코드</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="프로젝트코드" name="txtProjectCode" tag="24"  size=70 MAXLENGTH=100 value="<%=project_code%>"></TD>
							</TD>
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>제목</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<Select NAME="txtTitle" ALT="제목" STYLE="WIDTH: 150px" tag="22"><OPTION VALUE="" selected></OPTION></Select>
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>요약설명</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="요약설명" name="report_abbr" tag="21"  size=70 MAXLENGTH=100  ></TD>
						</TR>
				
						<TR>
							<TD CLASS="TD5" NOWRAP>작성자</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><input name="ins_person" alt="작성자" tag="22" MAXLENGTH=15   ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>등록일</TD>
							<TD CLASS="TD6" NOWRAP><input type=text name="insrt_dt" size=12>
							</TD>
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>내용</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><TEXTAREA class="tb4" alt="내용" cols=64 name="txtReportText" tag="22" rows=13 MAXLENGTH=1000  wrap=phsical><%= strReport_Text%></TEXTAREA></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>
								<INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbScript:vbAttachFile()' type=button value="첨부" id=button1 name=button1><BR>
								<INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick='vbScript:vbDeleteFile()' type=button value="목록에서 삭제" id=button2 name=button2></TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
								<SELECT  style="WIDTH: 470px" tag="21"  size=5 name=filelist multiple></SELECT></TEXTAREA></TD>
						</TR>
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD  HEIGHT=3 ></TD>
	</TR>
		
	<TR>
		<TD WIDTH=* ALIGN=RIGHT><input type=button name="OK" onclick="vbscript:CheckValid()" value="확인">&nbsp;&nbsp<input type=button name="Cancel" onclick="vbscript:FncClose()" value="닫기">&nbsp;</TD>
		<TD WIDTH=10>&nbsp;</TD>					
	</TR>	
	
</TABLE>
<table border=0 cellpadding="0" cellspacing="0" width="575" align=center>

</table>
</form>


<IFRAME  NAME="MyBizASP" STYLE="display: '';" WIDTH=100% HEIGHT=600 scrolling=yes></IFRAME>
<IFRAME NAME="FR_ATTWIZ" SRC="FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 ></IFRAME><BR>


</center>
</body>
</html>