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

<Script Language="VBScript">
'Option Explicit

Const BIZ_PGM_ID = "U2211PB1_KO441.asp"

Dim arrFileinf
Dim pStrFileInfo
Dim strMode
Dim arrTemp
Dim strBpCd, strDlvyNo
Dim popupParent
Dim arrParent

'Dim project_code

strMode  = "<%=strMode %>"
arrTemp  = "<%=arrTemp %>"
'project_code = "<%=project_code %>"
strBpCd		  = "<%=strBpCd%>"
strDlvyNo	  = "<%=strDlvyNo%>"

arrParent   = window.dialogArguments
Set popupParent = arrParent(0)

msgbox "여긴 왔나"
Function CheckValid()
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

    Call dbSave()
End Function

Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 

Private Sub Form_Load()
'Call ServerMesgBox("Form_Load - 249 " , vbInformation, I_MKSCRIPT)
'msgbox "a"
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, popupParent.gDateFormat, popupParent.gComNum1000, popupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N") 
	call SetDefaultVal()
'Call ServerMesgBox("Form_Load - 253 " , vbInformation, I_MKSCRIPT)
	If CStr(strMode) = "<%=CStr(UID_M0002)%>" Then	'글 등록이 아니면, 즉 글수정이면,
		Call FncQueryFileInfo()
		'frm1.use_dt.text  = UniConvDateAToB("<%=strUse_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=strIns_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	else
	
		'frm1.use_dt.text  = UniConvDateAToB("<%=dateadd("m",3,GetSvrDate)%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=GetSvrDate%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	End If
'Call ServerMesgBox("Form_Load - 263 " , vbInformation, I_MKSCRIPT)
	'Call InitComboBox

	frm1.txtTitle.value="<%=strDocument_No%>"
	frm1.Document_abbar.value="<%=strDocument_Abbr%>"
	frm1.ins_person.value="<%=strins_person%>"
	Call ggoOper.SetReqAttr(frm1.insrt_dt, "Q")
'Call ServerMesgBox("Form_Load - 270 " , vbInformation, I_MKSCRIPT)	
End Sub

''==========================================  2.2.6 InitComboBox()  ========================================
'' Name : InitComboBox()
'' Desc : Combo Display
''==========================================================================================================
'Sub InitComboBox()
'
'	'// 품의제목
'	Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD = " & FilterVar("SX006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'	Call SetCombo2(frm1.txtTitle , lgF0, lgF1, Chr(11))
'	
'End Sub

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
	
	'IntRetCD1= CommonQueryRs("REPORT_NM,REPORT_ID, REPORT_SIZE ","S_PRJ_REPORT_DTL_KO412", "project_code = '" & project_code & "' and report_no = '<%=strReport_No%>'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	IntRetCD1= CommonQueryRs("DOCUMENT_NM, DOCUMENT_ID, DOCUMENT_SIZE ","M_SCM_DOCUMENT_DTL_KO412", " BP_CD = '" & strBpCd & "' AND DLVY_NO = '" & strDlvyNo & "' and DOCUMENT_NO = '<%=strDocument_No%>'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

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
<input type=hidden name=BpCd value="<%=strBpCd%>">
<input type=hidden name=DlvyNo value="<%=strDlvyNo%>">
<input type=hidden name=DocumentNo value="<%=strDocument_No%>">
<Input type=hidden name=delTemp value="">

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
							<TD CLASS="TD5" NOWRAP>제목</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="제목" name="txtTitle" tag="22"  size=70 MAXLENGTH=100  >
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>요약설명</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="요약설명" name="Document_abbar" tag="21"  size=70 MAXLENGTH=100  ></TD>
						</TR>
				
						<TR>
							<TD CLASS="TD5" NOWRAP>등록자</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><input name="ins_person" alt="작성자" tag="22" MAXLENGTH=15   ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>등록일</TD>
							<TD CLASS="TD6" NOWRAP><input type=text name="insrt_dt" size=12>
							</TD>
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>내용</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><TEXTAREA class="tb4" alt="내용" cols=64 name="txtDocumentText" tag="22" rows=13 MAXLENGTH=1000  wrap=phsical><%= strReport_Text%></TEXTAREA></TD>
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