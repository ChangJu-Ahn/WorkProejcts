<% Option Explicit %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : 도면파일관리 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

'Option Explicit	
																'☜: indicates that All variables must be declared in advance
Dim IsOpenPop 
Dim arrReturn
Dim arrParent
Dim arrParam                         
Dim arrField
Dim PopupParent
                    
arrParent = window.dialogArguments

Set PopupParent = arrParent(0)

arrParam = arrParent(1)
arrField = arrParent(2)

top.document.title = PopupParent.gActivePRAspName

</SCRIPT>

<% 

Call LoadBasisGlobalInf()  
Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim strSystemFolder
Dim iTempPath
		
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
		

'========================================================================================================
' Name : GetSpecialFolder()     
' Desc : 
'========================================================================================================
Function GetSpecialFolder(iDx)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   GetSpecialFolder = pfile.GetSpecialFolder(CInt(iDx))   
   Set pfile = Nothing
End Function

'========================================================================================================
' Name : CreateFolder()     
' Desc : 
'========================================================================================================	
Function CreateFolder(iTempPath)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   Call pfile.CreateFolder(iTempPath)   
   Set pfile = Nothing
End Function
	
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<!--==========================================  1.1.2 공통 Include   =========================================-->

<SCRIPT Language="VBSCRIPT">

Const BIZ_PGM_ID  = "B82101pb4.asp"

Dim arrFileinf

'======================================================================================================
'	Name : FncClose()
'	Description : 
'======================================================================================================
Function FncClose()
	window.ReturnValue = false
	window.close()
	
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : 
'======================================================================================================
Private Sub Form_Load()		

	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, popupParent.gDateFormat, popupParent.gComNum1000, popupParent.gComNumDec)	
	Call ggoOper.LockField(Document, "N") 	
	Call InitVariables()
	Call SetDefaultVal()	
	Call QueryFileInfo()
	frm1.butS.disabled =true
	
End Sub


'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)	
'=======================================================================================================
Function InitVariables()

    Redim arrReturn(0)
    
    Self.Returnvalue = arrReturn    
    
    frm1.txtRet.value   = ""
    frm1.txtMode.value  = UID_M0001 'INSERT
        
End Function

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
    
    If arrParam(9) = "0" Then
       frm1.butB.disabled = False
       frm1.butS.disabled = False
       frm1.butD.disabled = False
    Else
       frm1.butB.disabled = True
       frm1.butS.disabled = True
       frm1.butD.disabled = True
    End If
   
    frm1.butC.disabled = False
    
    frm1.txtInternalCd.value = arrParam(0)
    frm1.txtItemCd.value     = arrParam(1)
    frm1.txtarReqNo.value     = arrParam(3)
   
    frm1.txtFileNm.focus 
End Sub

'======================================================================================================
'	Name : FncQueryFileInfo()
'	Description : 
'=======================================================================================================
Function QueryFileInfo()

	Dim IntRetCD
	Dim iLngRow
	Dim strData
	Dim strRet
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim ArrTmpF0,ArrTmpF1,ArrTmpF2, ArrTmpF3
	Dim strInternalCd,strReq_no
	
	'strInternalCd = frm1.txtInternalCd.value	
	strReq_no =frm1.txtarReqNo.value
	
	
	If strReq_no <> "" Then
	
	   IntRetCD= CommonQueryRs("FILE_NM, ID_FILE, FILE_SIZE, FILE_PATH ","B_CIS_DOCUMENT_FILE", "req_no = " & FilterVar(strReq_no, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	   ArrTmpF0 = split(lgF0,popupParent.gColSep)	
	   ArrTmpF1 = split(lgF1,popupParent.gColSep)	
	   ArrTmpF2 = split(lgF2,popupParent.gColSep)
	   ArrTmpF3 = split(lgF3,popupParent.gColSep)
	   
       strData = ""
	   strRet  = ""
	
	   For iLngRow = 0 To UBound(ArrTmpF0, 1) - 1
	
		   strData = "1" & ArrTmpF0(iLngRow) & "" & ArrTmpF2(iLngRow) & "" & ArrTmpF1(iLngRow) & "101344919970601092656N0NNFFY00YI"		
		   strRet = strRet & mid(strData,3,len(strData))
		
		   Call SetAttachFile(strData)
		
	   Next	
	
	   Call MakeFileInfoArray("1" & strRet)
	
	   If strData <> "" Then	  
	      
          frm1.txtFileNm.value = ArrTmpF0(0)	            
	      frm1.txtRet.value    = strData	      
	      frm1.txtSourceFilePath.value = ArrTmpF3(0) 
	      
	      frm1.txtMode.value = Popupparent.UID_M0002 'UPDATE
	      
	      If arrParam(9) = "0" Then
             frm1.butD.disabled = False
          End If
          
          frm1.butV.disabled = False
          frm1.butW.disabled = False
       
       Else
          frm1.butV.disabled = True
          frm1.butW.disabled = True
          frm1.txtMode.value = Popupparent.UID_M0001 'INSERT
	   End If 
	   
	End If    
		
End Function


'======================================================================================================
'	Name : FncSave()
'	Description : 
'=======================================================================================================
Function FncSave()

	If UCASE(frm1.txtblnFileAttached.value) = "TRUE" Then
	   
	   Call DbSave()
	   
	End If
	
End Function


'======================================================================================================
'	Name : DbSave()
'	Description : 
'=======================================================================================================
Function DbSave()

	Call MakeFileInfoString(arrFileinf)
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
	Call QueryFileInfo()
	
	frm1.butS.disabled = True  '저장버튼 비활성화 
	frm1.butD.disabled = False '삭제버튼 활성화 
	
End Function

'======================================================================================================
'	Name : DbSaveOk()
'	Description : 
'=======================================================================================================
Function DbSaveOk()

	window.ReturnValue = True
	Self.Close()
	
End Function

'======================================================================================================
'	Name : FncDelete()
'	Description : 
'=======================================================================================================
Function FncDelete()

	If UCASE(frm1.txtblnFileAttached.value) = "TRUE" Then
	
	   frm1.txtMode.value = Popupparent.UID_M0003 'DELETE
	
	   Call DbSave()
	   
	   Call ggoOper.ClearField(Document, "A")
	   
	   frm1.txtInternalCd.value = arrParam(0)
       frm1.txtItemCd.value     = arrParam(1)
	   
	   frm1.butD.disabled = True '삭제버튼 비활성화 
	   frm1.butV.disabled = True '파일보기 비활성화 
       frm1.butW.disabled = True '파일저장 비활성화 
	   
	End If
	
End Function

'======================================================================================================
'	Name : vbAttachFile()
'	Description : 
'=======================================================================================================
Function vbAttachFile()

	Dim strRet
		
	strRet = AttachFile()
	
	If Len(strRet) > 3  then
	
	   Call MakeFileInfoArray(strRet)
	      
	   frm1.txtRet.value  = strRet
	   frm1.butS.disabled = False  '저장버튼 활성화 
	   
	End if
	
End Function

'======================================================================================================
'	Name : MakeFileInfoArray()
'	Description : 
'=======================================================================================================
Function MakeFileInfoArray(strRet)

	Dim arrTemp,arrTemp2
	Dim i,j
	Dim iFileCount
	Dim iCurrentSize
	
	
	arrFileinf = ""
	
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

End Function

'======================================================================================================
'	Name : MakeFileInfoString()
'	Description : 
'=======================================================================================================
Function MakeFileInfoString(prArrTemp)

	Dim i
	Dim strFileInfo

	Const C_FileName  = 2
	Const C_FileId    = 4	
	Const C_FileSize  = 9
	Const C_FileCDate = 10
	
	For i = 1 To UBound(prArrTemp,2)
	
		strFileInfo = strFileInfo & prArrTemp(C_FileName, i)  & popupParent.gColSep
		strFileInfo = strFileInfo & prArrTemp(C_FileId, i)    & popupParent.gColSep
		strFileInfo = strFileInfo & prArrTemp(C_FileSize, i)  & popupParent.gColSep
		strFileInfo = strFileInfo & prArrTemp(C_FileCDate, i) & popupParent.gRowSep
					
	Next
	
	frm1.txtFileinf.value = strFileInfo

End Function

'======================================================================================================
'	Name : vbDeleteFile()
'	Description : 
'=======================================================================================================
Function vbDeleteFile

	Dim iDx
	Dim strTemp
	
	If frm1.txtFileList.length > 0 Then
		
	   For iDx = CInt(frm1.txtFileList.length) - 1 To 0 Step -1
		   frm1.txtFileList.remove(iDx)
		   RearrangeArray(iDx + 1)			
	   Next
	   
       frm1.txtFileNm.value = ""
              
	Else		
		Exit Function
	End If
	
End Function

'======================================================================================================
'	Name : FetchWebSvrIp()
'	Description : 
'=======================================================================================================
Function FetchWebSvrIp()	

	Dim gHttpWebSvrIPURL
	
	gHttpWebSvrIPURL =  "http://<%= request.servervariables("server_name") %>"	
	FetchWebSvrIp = Split(gHttpWebSvrIPURL, "/")(2)
	
End Function

'======================================================================================================
'	Name : FncViewFile()
'	Description : 
'=======================================================================================================
Function FncFileView()
        
	 Call vbFileView("F")
	
End Function

'======================================================================================================
'	Name : FncSaveFile()
'	Description : 
'=======================================================================================================
Function FncFileSave()
    
    Call vbFileView("W")
	
End Function

'======================================================================================================
'	Name : VbViewFile()
'	Description : 
'=======================================================================================================
Function vbFileView(asMod)
    
    Dim strRet
    Dim arrTemp , arrTemp2
    Dim strLocalPath
    Dim strServerPath
    Dim iPos
        
    strRet = frm1.txtRet.value
    
    arrTemp  = Split(strRet, chr(31))
    arrTemp2 = Split(arrTemp(1), chr(29))
    
    strLocalPath   = "<%=Request.Cookies("unierp")("gTempDirForFileUpload")%>" & arrTemp2(3)     
    strLocalPath   = Replace(strLocalPath,"/", "\")    
    strServerPath  = frm1.txtSourceFilePath.value
       
    iPos           = InStrRev(strRet, arrTemp2(3), -1)    
    strRet         = Mid(strRet , 1, iPos - 1) & strLocalPath & "" & strServerPath & Mid(strRet , iPos , Len(strRet))
    
    frm1.txtViewRet.value        = strRet
    frm1.txtTargetPath.value     = "<%=Request.Cookies("unierp")("gTempDirForFileUpload")%>"
    frm1.txtTargetFileName.value = arrTemp2(3) 
    
	frm1.txtFileMod.value = asMod
		
	Call ExecMyBizASP(frm1, "B82101pb5.asp")
	
	//Call ViewFile(asMod, strRet )
	
End Function

'=======================================================================================================

</SCRIPT>

<SCRIPT language=JavaSCRIPT>

//======================================================================================================
//	Name : AttachFile()
//	Description : 
//=======================================================================================================

function AttachFile(){

	var sRet;
	var optStr;
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
		
		if(iFileinf.length > 1){
			optStr = new Option(iFileinf[1],sRet,true);
			//파일은 하나만 선택하는것으로 하기 때문에... 무조건 [0]으로 하고 바로 빠져나온다.
			//document.frm1.txtFileList.options[document.frm1.txtFileList.length] = optStr;
			document.frm1.txtFileList.options[0] = optStr;
			document.frm1.txtFileNm.value = iFileinf[1];
			return(sRet);
		}
	}	
	return(sRet);
}


//======================================================================================================
//	Name : SetAttachFile()
//	Description : 
//=======================================================================================================

function SetAttachFile(sRet){

	var arrTemp = sRet.split(String.fromCharCode(31));
	
	for(var i = 0; i < arrTemp.length; i++){		
		var iFileinf = arrTemp[i].split(String.fromCharCode(29));
		
		if(iFileinf.length > 2){
			optStr = new Option(iFileinf[1],sRet,true);				
			document.frm1.txtFileList.options[document.frm1.txtFileList.length] = optStr;
			document.frm1.txtFileNm.value = iFileinf[1];
		}
	}		
}


//======================================================================================================
//	Name : ViewFile()
//	Description : 
//=======================================================================================================
function ViewFile(sMode, sRet){
	
	var strWebSvrIp;
	
	document.FR_ATTWIZ.SetLanguage('K');	
	document.FR_ATTWIZ.SetModUpload();
	document.FR_ATTWIZ.SetServerAutoDelete(1);
	document.FR_ATTWIZ.SetFileUIMode(1);
	document.FR_ATTWIZ.SetServerOption(0,0);	
    document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	document.FR_ATTWIZ.ViewFile(sMode, sRet);
}	


</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<BODY SCROLL=NO TABINDEX="-1">

<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=10>
			<FIELDSET CLASS="CLSFLD">
			    <BR>
			    <LEGEND>도면파일선택</LEGEND>
			    <BR>
				<TABLE WIDTH=100% CELLSPACING=0>
				   <TD HEIGHT=30>
				      <TD WIDTH=5 NOWRAP></TD>
				      <TD WIDTH=100% ALIGN=midle><INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 19px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:vbAttachFile()' type=button value='찾아보기..' id=butB name=butB>
				                                 <INPUT NAME="txtFileNm" ALT="파일명" TYPE="Text" SiZE=35 MAXLENGTH=100  tag="24XXXX"></TD>
                      <TD WIDTH=5>&nbsp;</TD>    
                   </TD>                   
				</TABLE>				
				<TABLE WIDTH=100% CELLSPACING=0>
				   <TD HEIGHT=30>
				      <TD WIDTH=5 NOWRAP></TD>
				      <TD WIDTH=100% ALIGN=RIGHT><INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:FncFileView()' type=button value='파일보기' id=butV name=butV>
                                                 <INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:FncFileSave()' type=button value='파일저장' id=butW name=butW></TD>
                      <TD WIDTH=5>&nbsp;</TD>    
                   </TD>                   
				</TABLE>                         								
				<BR>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
	   <TD HEIGHT=10>
		    <TABLE CLASS="basicTB" CELLSPACING=0>
			    <TR>
				   <TD HEIGHT=50>
				      <TD WIDTH=10% NOWRAP></TD>
				      <TD WIDTH=90% ALIGN=RIGHT><INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:FncSave()'     type=button value='저장' id=butS name=butS>
				                                <INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:FncDelete()'   type=button value='삭제' id=butD name=butD>
				                                <INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbSCRIPT:FncClose()'    type=button value='닫기' id=butC name=butC></TD>
                      <TD WIDTH=10>&nbsp;</TD>
                   </TD>
			   </TR>
		   </TABLE>
	   </TD>
	</TR>
</TABLE>

<INPUT type=hidden name=txtMode            value="">
<INPUT type=hidden name=txtblnFileAttached value="">
<INPUT type=hidden name=txtInternalCd      value="">
<INPUT type=hidden name=txtItemCd          value="">
<INPUT type=hidden name=txtFileinf         value="">
<INPUT type=hidden name=txtarReqNo         value="">
<SELECT type=hidden name=txtFileList multiple></SELECT>

<INPUT type=hidden name=txtFileMod         value="">
<INPUT type=hidden name=txtRet             value="">
<INPUT type=hidden name=txtViewRet         value="">
<INPUT type=hidden name=txtSourceFilePath  value="">
<INPUT type=hidden name=txtTargetPath      value="">
<INPUT type=hidden name=txtTargetFileName  value="">

</FORM>
<IFRAME  NAME="MyBizASP" STYLE="display: '';" WIDTH=100% HEIGHT=300 scrolling=yes></IFRAME>

<IFRAME NAME="FR_ATTWIZ" SRC="../../Notice/FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 ></IFRAME>

</BODY>
</HTML>