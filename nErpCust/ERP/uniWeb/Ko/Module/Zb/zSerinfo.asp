<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : System Information
'*  3. Program ID           : zSerinfo.asp
'*  4. Program Name         : System Information
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/11/13
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Seung jin
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncServer.asp" -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
'=========================================================================================================
Sub Form_Load()

    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("10000000000000")                                         
     dim IntRetCD
     Dim MaxVal
     Dim MinVal
     Dim MemCfg
     IntRetCD=CommonQueryRs("u.value","master.dbo.spt_values v  left outer join master.dbo.sysconfigures  c  on v.number = c.config left outer join master.dbo.syscurconfigs  u  on v.number = u.config","name='min server memory (MB)'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
     lgf0 = split(lgf0,Chr(11))
     MinVal = lgf0(0)
     IntRetCD=CommonQueryRs("u.value","master.dbo.spt_values v  left outer join master.dbo.sysconfigures  c  on v.number = c.config left outer join master.dbo.syscurconfigs  u  on v.number = u.config","name='max server memory (MB)'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
     lgf0 = split(lgf0,Chr(11))
     MaxVal = lgf0(0)
     
     if MinVal = "2147483647" then
      MinVal = "255"
     end if
     
     if MaxVal = "2147483647" then
      MaxVal = "255"
     end if
     
     if MinVal = MaxVal then
      MemCfg = " 고정메모리 구성"
     else
      MemCfg = " 동적메모리 구성"
     end if
     
     'msgbox "Minimum : " & MinVal & " / Maximum :" & MaxVal
     DBMaxMinMemVal.innerHTML= "Minimum : " & MinVal & " / Maximum :" & MaxVal & MemCfg

End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
   FncExit = True																		'☜: Processing is OK
End Function

</SCRIPT>
<%
On Error Resume Next
 
Dim objSvr

Dim strWebAppServerIP
Dim strSvrConfigure
Dim strGetDBIP

'OS 관련 
Dim strGetOSProduct
Dim strGetOSVersion
Dim strGetOSVersionType
Dim strGetOSBuild
Dim strGetOSProductID 
Dim strGetOSServicePack
Dim strGetHotFix
Dim strGetHotFixEle
Dim HotFixCnt
Dim strHotFixElement

'DB Server 관련 
Dim strGetDBInstallDir
Dim strGetDBVersion
Dim strGetDBLang

'Memory 관련 
Dim strGetLoadMemInfo
Dim strGetTotalMemInfo
Dim strGetAvailMemInfo
Dim strGetMemInfoEle

'CPU 관련 
Dim strGetCPUName
Dim strGetIdentifier


'DBAgent 관련 
Dim iPath
Dim LoginVD

'uniERP 관련 
Dim Lang
Dim LangDesc
Dim uniVersion
Dim verSpec
Dim iniFileDir
Dim strGetInstallDir
Dim GetSetupModuleCnt
Dim strCompany
Dim CompanyName
Dim Msgnodata

'DriveSpace 관련 
Dim strGetDriveSpace
Dim strGetDriveFreeSpace

'포맷정보 
Const LOCALE_SDECIMAL   = &HE     'decimal separator
Const LOCALE_STHOUSAND  = &HF     'thousand separator
Const LOCALE_SSHORTDATE = &H1F    'short date format string
Dim strSDECIMAL
Dim strSTHOUSAND
Dim strSSHORTDATE


Set objSvr               = CreateObject("ServerInfoControl.clsSrvInfo") 
    
    Select Case gLang
        Case "KO"
                LangDesc  = "Korean"
        Case "CN"
                LangDesc  = "Chinese"
        Case "EN"
                LangDesc  = "English"
        Case "JA"
                LangDesc  = "Japanese"
        Case Else         
                LangDesc  = "언어에 대한 정보가 없습니다."
    End Select 

Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "S", "", strGetOSProduct)
Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentVersion", "S", "", strGetOSVersion)
Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId", "S", "", strGetOSProductID)
Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber", "S", "", strGetOSBuild)
Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion", "S", "", strGetOSServicePack)
strGetHotFix = objSvr.FncReadKeyList("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix")
   
   if strGetHotFix <> "" then
       strGetHotFix = UCase(strGetHotFix)
       HotFixCnt = 0
       strGetHotFixEle = Split(strGetHotFix,Chr(11))  ' Chr(11) 특수기호 
       strGetHotFix = ""
       
       for each strHotFixElement in strGetHotFixEle
         strGetHotFix = strGetHotFix & strHotFixElement & " "
         
         HotFixCnt = HotFixCnt + 1
         
            if HotFixCnt Mod 3 = 0 then
            strGetHotFix = strGetHotFix & "</BR>"
        End If
       next 
   end if

   if Trim(strGetOSProduct) = "Microsoft Windows 2000" then

     Call objSvr.FncRegReadValue("HKLM", "SYSTEM\CurrentControlSet\Control\ProductOptions", "ProductType", "S", "", strGetOSVersionType)               
               
     Select Case UCase(strGetOSVersionType)
        Case "WINNT"
            strGetOSProduct = strGetOSProduct + " Professional"
        Case "LANMANNT"
            strGetOSProduct = strGetOSProduct + " Server"
        Case "SERVERNT"
            strGetOSProduct = strGetOSProduct + " Server"
     End Select
   
   end if

   If Trim(strGetOSServicePack) = "" Then strGetOSServicePack = "No Setting"
      Call objSvr.FncRegReadValue("HKLM", "Hardware\Description\System\CentralProcessor\0", "ProcessorNameString", "S", "", strGetCPUName)
      Call objSvr.FncRegReadValue("HKLM", "Hardware\Description\System\CentralProcessor\0", "Identifier", "S", "", strGetIdentifier)
      Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\MSSQLServer\MSSQLServer\CurrentVersion", "CurrentVersion", "S", "", strGetDBVersion)
      Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\MSSQLServer\MSSQLServer\CurrentVersion", "Language", "D", "", strGetDBLang)

      Select Case Hex(strGetDBLang)
        Case "409"
                strGetDBLang = "English"
        Case "412"
                strGetDBLang = "Korean"
        Case "411"
                strGetDBLang = "Japanese"
        Case "40e"
                strGetDBLang = "Hungarian"
        Case "407"
                strGetDBLang = "German"
        Case "804","404"
                strGetDBLang = "Chinese"
      End Select 
      
'Memory info
Call objSvr.FncMemoryInfo(strGetMemInfoEle) 
strGetMemInfoEle = Split(strGetMemInfoEle,Chr(11))  ' Chr(11) 특수기호 
strGetLoadMemInfo = strGetMemInfoEle(0) '현재 사용중인 메모리 
strGetTotalMemInfo = strGetMemInfoEle(1) ' Total Memory
strGetAvailMemInfo = strGetMemInfoEle(2) 'Availible Memory

'version info

Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\SAMSUNG_SDS\unisetup", "version", "S", "", uniVersion)
If Trim(uniVersion) = "" Then
   Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\unisetup", "version", "S", "", uniVersion)
End If


        Select Case verSpec
        Case "20", "25"
        
          Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\unisetup\Web", "installdir", "S", "", strGetInstallDir)
          iniFileDir = strGetInstallDir & "uniWeb\uniDefaultport.ini"
          
        Case "27"
        
          Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\SAMSUNG_SDS\unisetup", "installdir_webapp", "S", "", strGetInstallDir)
          iniFileDir = strGetInstallDir & "uniWeb\uniSystemInfo.ini"
        
        Case Else
          
          Call objSvr.FncRegReadValue("HKLM", "SOFTWARE\SAMSUNG_SDS\unisetup", "installdir_webapp", "S", "", strGetInstallDir)
          iniFileDir = strGetInstallDir & "uniWeb\uniSystemInfo.ini"
          
        End Select
    
   'khy
    iPath = Split(Request.ServerVariables("PATH_INFO"),"/")
    LoginVD = UCase(iPath(UBound(iPath) - 4))
    
    Dim IntRetCD, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    IntRetCD = CommonQueryRs("CO_FULL_NM","B_COMPANY","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)		
    lgF0 = Replace(lgF0, Chr(11), vbTab)    
    lgF0 = Replace(lgF0, " ","")
    CompanyName = lgF0   
       
    Call objSvr.FncDiskSpace(strGetInstallDir, strGetDriveSpace, strGetDriveFreeSpace)
    Call objSvr.FncLocalRegionOpt(LOCALE_SDECIMAL,strSDECIMAL)
    Call objSvr.FncLocalRegionOpt(LOCALE_STHOUSAND,strSTHOUSAND)
    Call objSvr.FncLocalRegionOpt(LOCALE_SSHORTDATE,strSSHORTDATE)

    '서버구성 체크 
    strWebAppServerIP = request.servervariables("server_name")
    If Trim(strWebAppServerIP) = Trim(strGetDBIP) then
       strSvrConfigure = "Web/App/DB 통합 서버(" & strWebAppServerIP & ")"
    Else
       strSvrConfigure = "Web/App/DB 분리 서버(AppWeb:" & strWebAppServerIP & " , DB:" & strGetDBIP & ")"
    End if

    Set objSvr = nothing
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        
</HEAD>
<BODY TABINDEX="-1" SCROLL="Yes">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
                                <td background="../../../CShared/image/table/seltab_up_left.gif"  width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>서버정보</font></td>
                                <td background="../../../CShared/image/table/seltab_up_right.gif" align="right"  width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=*>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% CLASS="Tab11" valign=top>
            <TABLE width=100% valign=top border = 0>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <LEGEND>서버 구성형태</LEGEND>
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>서버 구성</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strSvrConfigure%></TD>    
                                    <TD CLASS=TD6 NOWRAP></TD>
                                    <TD CLASS=TD6 NOWRAP></TD>
                                </TR>
                            </TABLE>
                        </FIELDSET>
                        <BR>
                        <FIELDSET CLASS="CLSFLD">
                            <LEGEND>OS 정보</LEGEND>
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>제품명</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetOSProduct%></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP>버젼</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetOSVersion%></TD>                        
                                </TR>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>Build명</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetOSBuild%></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP>서비스팩 정보</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetOSServicePack%></TD>                        
                                </TR>
                                <TR>

                                    <TD CLASS=TD5 NOWRAP vAlign=TOP>HotFix 설치 정보</TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP><%=strGetHotFix%></TD>
                                    
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP></TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP></TD>                            
                                </TR>
                            </TABLE>
                        </FIELDSET>
                        <BR>
                        <FIELDSET CLASS="CLSFLD">
                            <LEGEND>MS-SQL 정보</LEGEND>
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP>버젼</TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP><%=strGetDBVersion%></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP>언어</TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP><%=strGetDBLang%></TD>                        
                                </TR>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP>메모리 용량 (MB)</TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP id=DBMaxMinMemVal></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP></TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP></TD>                        
                                </TR>
                            </TABLE>
                        </FIELDSET>
                        <BR>
                        <FIELDSET>
                            <LEGEND>CPU / Memory 정보</LEGEND>
                                <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>CPU 이름</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetCPUName%></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP>CPU 식별자</TD>
                                    <TD CLASS=TD6 NOWRAP><%=strGetIdentifier%></TD>                        
                                </TR>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>물리적 메모리 용량</TD>
                                    <TD CLASS=TD6><%=strGetLoadMemInfo & "% in Use (Total:"& formatnumber(strGetTotalMemInfo,0) & "KB / Free:" & formatnumber(strGetAvailMemInfo,0) & "KB)"%></TD>    
                                    <TD CLASS=TD5 NOWRAP vAlign=TOP>디스크 공간</TD>
                                    <TD CLASS=TD6 NOWRAP vAlign=TOP><%=strGetDriveSpace%>MB (Free : <%=strGetDriveFreeSpace%>MB)</TD>    
                                </TR>
                            </TABLE>
                        </FIELDSET>
                        <BR>
                        <FIELDSET>
                            <LEGEND>포맷 정보</LEGEND>
                                <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD>
                                    
                                        <FIELDSET>
                                            <LEGEND>OS 포맷 정보</LEGEND>
                                                <TABLE <%=LR_SPACE_TYPE_40%>>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>소수점 구분자</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=strSDECIMAL%></TD>                        
                                                </TR>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>1000 단위 구분자</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=strSTHOUSAND%></TD>                        
                                                </TR>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>날짜 형식</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=strSSHORTDATE%></TD>                        
                                                </TR>
                                            </TABLE>
                                        </FIELDSET>
                                        
                                    </TD>
                                    <TD>
                                    
                                        <FIELDSET>
                                            <LEGEND>uniERP 포맷 정보</LEGEND>
                                                <TABLE <%=LR_SPACE_TYPE_40%>>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>소수점 구분자</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=gComNumDec%></TD>                        
                                                </TR>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>1000 단위 구분자</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=gComNum1000%></TD>                        
                                                </TR>
                                                <TR>
                                                    <TD CLASS=TD5 NOWRAP>날짜 형식</TD>
                                                    <TD CLASS=TD6 NOWRAP><%=gDateFormat%></TD>                        
                                                </TR>
                                            </TABLE>
                                        </FIELDSET>
                                        
                                    </TD>                        
                                </TR>
                            </TABLE>
                        </FIELDSET>
                        <BR>
                    </TD>
                </TR>

            </TABLE>
        </TD>
    </TR>
</TABLE>
</body>
</html>

