<%

Sub CheckServer()

      Dim xDom
      Dim iStub
      
      Set iStub = Server.CreateObject("uniStub.CX01")
      
      Set xDom = Server.CreateObject(gMSXMLDOMDocument)		
      
      xDom.async = False 
      
      xDom.loadXML (iStub.IGetNotifierData)
      
      Set iStub = Nothing
      
      If xDom.selectSingleNode("/DATA/STATUS").Text = "MES" Then
         If xDom.selectSingleNode("/DATA/MKIND").Text = "S" Then
            If CInt(xDom.selectSingleNode("/DATA/TIMEOUT").Text) <= 1 Then
               Set xDom = Nothing
               Response.Redirect "Notifier.asp"
            End If
         End If
      End If
      
      Set xDom = Nothing

End Sub  


Sub WriteDownComScript()

   Response.Write "<DIV  style='display:none;'>" & vbCrLf
   Response.Write "<OBJECT ID=DOMDocument401   CLASSID=CLSID:88d969c0-f192-11d4-a65f-0040963251e5 codebase=""control/msxml4.cab#version=4,10,9404,0""></object>"   & vbCrLf
   Response.Write "<OBJECT                     CLASSID=CLSID:E76EE1AE-B057-47FC-ABBE-AC109ED0979A codebase=""control/unicencrypt.cab#version=1,0,0,1""></object>"   & vbCrLf
   Response.Write "<OBJECT ID=ConnectorControl CLASSID=CLSID:EBCF5077-3931-11D4-8D9D-00AA005703C2 codebase=""control/uni2kCommon.CAB#version=1,0,0,25""></OBJECT>" & vbCrLf
   Response.Write "<OBJECT ID=HndRegCls        CLASSID=CLSID:333EE8B1-4927-11D4-BC29-00AA00BFC320  ></OBJECT>" & vbCrLf
   Response.Write "</DIV>" & vbCrLf

End Sub

Sub WriteDownESSComScript()

   Response.Write "<DIV  style='display:none;'>" & vbCrLf
   Response.Write "<OBJECT ID=DOMDocument401   CLASSID=CLSID:88d969c0-f192-11d4-a65f-0040963251e5  codebase=""./ESSControl/msxml4.cab#version=4,10,9404,0""></object>"   & vbCrLf
   Response.Write "<OBJECT                     CLASSID=CLSID:E76EE1AE-B057-47FC-ABBE-AC109ED0979A  codebase=""./ESSControl/unicencrypt.cab#version=1,0,0,1"" id=CEncrypt_Class1></object>"   & vbCrLf
   Response.Write "<OBJECT ID=ConnectorControl CLASSID=CLSID:EBCF5077-3931-11D4-8D9D-00AA005703C2  codebase=""./ESSControl/uni2kCommon.CAB#version=1,0,0,25""></OBJECT>" & vbCrLf
   Response.Write "<OBJECT ID=HndRegCls        CLASSID=CLSID:333EE8B1-4927-11D4-BC29-00AA00BFC320  ></OBJECT>" & vbCrLf
   Response.Write "</DIV>" & vbCrLf

End Sub

Sub WriteWindow_onLoad()

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Sub Window_onLoad() " & vbCrLf

    Response.Write "    Dim strURLLang       " & vbCrLf
    Response.Write "    Dim strURLLangUserID " & vbCrLf
    Response.Write "    Dim iTemp            " & vbCrLf
    
    Response.Write "    iTemp = Split(document.location.href,""/"")  " & vbCrLf
    
    Response.Write "    frmLogin.initp.value = iTemp(UBound(iTemp))  " & vbCrLf
    
    Response.Write "    gLogo = """ & Request.Cookies("unierp")("gLogoName") & """" & vbCrLf
    
    Response.Write "    If """ & gSSO & """ = ""Y"" Then " & vbCrLf
    
    Response.Write "       MouseWindow.location.href= ""./inc/SSO.htm""" & vbCrLf
    Response.Write "       MousePT.style.visibility = ""visible"" " & vbCrLf

    Response.Write "	   strURLLang =  LCase(frmLogin.txtURL.value)  & ""/"" & LCase(frmLogin.txtLangCdForURL.value)  " & vbCrLf
    Response.Write "	   strURLLangUserID = ""["" & strURLLang & ""][UID:""" & " & """ & Request("uid") & """ & "  & """]"" " & vbCrLf 
	   
    Response.Write "	   If CallDirectPage(strURLLangUserID) = False Then   " & vbCrLf
    Response.Write "          txtUsrIdview.value = """ & Request("uid") & """ " & vbCrLf
    Response.Write "          txtpwdview.value   = """ & Request("pwd") & """ " & vbCrLf
    Response.Write "          Call OKProcess() " & vbCrLf
    Response.Write "       End If    " & vbCrLf

    Response.Write "    Else    " & vbCrLf

    Response.Write "       txtUsrIdview.focus " & vbCrLf
       
    Response.Write "    End If    " & vbCrLf

    Response.Write "	If """ & Request.Cookies("gSaveID") & """  <> """"  And """ & gSSO & """ = """"  Then  " & vbCrLf
    Response.Write "		txtUsrIdview.value= """ & Request.Cookies("gSaveID") & """" & vbCrLf
    Response.Write "		IDSaveview.checked  = True		 " & vbCrLf
    Response.Write "		txtpwdview.focus " & vbCrLf
    Response.Write "	End If " & vbCrLf
	
	
    Response.Write "    If gMC = ""Y"" Then            " & vbCrLf
    Response.Write "       If gDebugMode = ""Y"" Then  " & vbCrLf
    Response.Write "          Call LoadCompInfoD()     " & vbCrLf
    Response.Write "       Else                        " & vbCrLf
    Response.Write "          Call LoadCompInfoG()     " & vbCrLf
    Response.Write "       End If                      " & vbCrLf
    Response.Write "    End If                         " & vbCrLf
	
    Response.Write "End Sub " & vbCrLf

End Sub


Sub WriteESSWindow_onLoad()

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Sub Window_onLoad() " & vbCrLf

    Response.Write "    Dim strURLLang       " & vbCrLf
    Response.Write "    Dim strURLLangUserID " & vbCrLf
    Response.Write "    Dim iTemp            " & vbCrLf
    
    Response.Write "    iTemp = Split(document.location.href,""/"")  " & vbCrLf
    
    Response.Write "    frmLogin.initp.value = iTemp(UBound(iTemp))  " & vbCrLf
    

    Response.Write "       txtUsrIdview.focus " & vbCrLf
       
    Response.Write "	If """ & Request.Cookies("gESSSaveID") & """  <> """"  Then  " & vbCrLf
    Response.Write "		txtUsrIdview.value= """ & Request.Cookies("gESSSaveID") & """" & vbCrLf
    Response.Write "		IDSaveview.checked  = True		 " & vbCrLf
    Response.Write "		txtpwdview.focus " & vbCrLf
    Response.Write "	End If " & vbCrLf
	
	
    Response.Write "End Sub " & vbCrLf

End Sub

Sub WriteCallDirectPage()

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Function CallDirectPage(ByVal strURLLangUserID)          " & vbCrLf
    Response.Write "    Dim objConn                                          " & vbCrLf
    
    Response.Write "    CallDirectPage = False                               " & vbCrLf 

    Response.Write "    Set objConn = CreateObject(""uniConnector.cGlobal"") " & vbCrLf
    
    Response.Write "    If Err.number <> 0 Then " & vbCrLf
    Response.Write "       MsgBox ""uniConnector 로드중 오류가 발생했습니다."" & vbCrLf &  vbCrLf & ""원인 : ["" & Err.number & ""] "" & Err.Description ,vbExclamation,""" & Request.Cookies("unierp")("gLogoName") & """" & vbCrLf
    Response.Write "       Exit Function    " & vbCrLf
    Response.Write "    End If " & vbCrLf

    Response.Write "    If objConn.ExistsURL (strURLLangUserID) = True Then " & vbCrLf
    Response.Write "       objConn.InitURL (strURLLangUserID)               " & vbCrLf
    Response.Write "       top.location = ""./SessionTrans.Asp?DPC=Y&DPCP=""" & " & """ & Request("DPCP") & """ & " & """&arg=""" & " & """ & Request("arg") & """ & " & """&"" & objConn.GetAspPostString() " & vbCrLf
    Response.Write "       CallDirectPage = True " & vbCrLf
    Response.Write "    End If    " & vbCrLf

    Response.Write "    Set objConn = Nothing    " & vbCrLf
    Response.Write "End Function     " & vbCrLf

End Sub


Sub WriteCheckKey()

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write " Function CheckKey() " & vbCrLf
    
    Response.Write "     If window.event.keyCode = 13 Then " & vbCrLf
    Response.Write "        Call OKProcess()               " & vbCrLf
    Response.Write "     End If                            " & vbCrLf
    Response.Write " 	                                   " & vbCrLf
    Response.Write " End Function                          " & vbCrLf

End Sub

Sub WriteOKProcess()

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Function OKProcess() " & vbCrLf

    Response.Write "	If CheckIDPWD = False Then " & vbCrLf
    Response.Write "	   Exit Function           " & vbCrLf
    Response.Write "	End If                     " & vbCrLf
	
    Response.Write "	If gMC = ""Y"" Then        " & vbCrLf
    Response.Write "       If gDebugMode = ""Y"" Then " & vbCrLf
    Response.Write "          frmLogin.txtVD.value   = CompanyList.value & ""_"" & DbList.value " & vbCrLf
    Response.Write "        Else                                         " & vbCrLf
    Response.Write "          frmLogin.txtVD.value   = CompanyList.value " & vbCrLf
    Response.Write "       End If                                        " & vbCrLf
    Response.Write "    End If                                           " & vbCrLf
    	
    Response.Write "	frmLogin.Action = ""UniLoginProcess.ASP""        " & vbCrLf
    Response.Write "	frmLogin.submit                                  " & vbCrLf

    Response.Write "End function                                         " & vbCrLf
    
End Sub

Sub WriteESSOKProcess(ByVal pASP)

    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Function OKProcess() " & vbCrLf

    Response.Write "	If CheckIDPWD = False Then " & vbCrLf
    Response.Write "	   Exit Function           " & vbCrLf
    Response.Write "	End If                     " & vbCrLf
	
    If pASP	= "S" Then
       Response.Write "	frmLogin.Action = ""ESSLoginprocess.asp?SAPP=ESS"" " & vbCrLf
    Else
       Response.Write "	frmLogin.Action = ""./module/e1/elogin_ok.asp?SAPP=ESS"" " & vbCrLf
    End If   
    Response.Write "	frmLogin.submit                                    " & vbCrLf

    Response.Write "End function                                           " & vbCrLf
    
End Sub

Sub WriteHTTPFormString()

    Response.Write "<INPUT TYPE=hidden NAME=txtUsrId                             > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtPwd                               > " & vbCrLf			    

    Response.Write "<INPUT TYPE=hidden NAME=txtURL                               > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtVD                                > " & vbCrLf

    Response.Write "<INPUT TYPE=hidden NAME=ChkIdSave                            > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=blnFlagSingleLogOnOrNot value=uniSIMS> " & vbCrLf			    
    Response.Write "<INPUT TYPE=hidden NAME=txtCanBeDebug Value= 2               > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtCnt value=0                       > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtHttpWebSvrIPURL                   > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtILVL      Value= RC               > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtLang                              > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtLangCdForURL                      > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=IDSave                               > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=initp                                > " & vbCrLf
    
    Response.Write "<INPUT TYPE=hidden NAME=txtClientNum1000                     > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtClientNumDec                      > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtClientDateFormat                  > " & vbCrLf
    Response.Write "<INPUT TYPE=hidden NAME=txtClientDateSeperator               > " & vbCrLf

End Sub
Sub WriteConfiguration()
    Response.Write "'================================================================================================================== " & vbCrLf
    Response.Write "Function Configuration()  " & vbCrLf
    Response.Write "                          " & vbCrLf
    Response.Write "	Dim strRet            " & vbCrLf
    Response.Write "                          " & vbCrLf
    Response.Write "    strRet = window.showModalDialog(""uniConfig.asp"",,""dialogWidth=20.3em; dialogHeight=20.5em; center: Yes; help: No; resizable: No; status: No;"") " & vbCrLf
    Response.Write "                                                             " & vbCrLf
    Response.Write "	strLang =  """ & Request.Cookies("unierp")("gLang") & """" & vbCrLf
    Response.Write "                                                      " & vbCrLf
    Response.Write "    If strRet = ""KO"" and  strLang <> ""KO"" Then    " & vbCrLf
    Response.Write "       window.location.href = ""../ko/unilogin.asp""  " & vbCrLf
    Response.Write "    Elseif strRet = ""EN"" and strLang <> ""EN"" Then " & vbCrLf
    Response.Write "       window.location.href = ""../en/unilogin.asp""  " & vbCrLf
    Response.Write "    Elseif strRet = ""CN"" and strLang <> ""CN"" Then " & vbCrLf
    Response.Write "       window.location.href = ""../cn/unilogin.asp""  " & vbCrLf
    Response.Write "    End if                   " & vbCrLf
    Response.Write "                             " & vbCrLf
    Response.Write "    Call Window_OnLoad()	 " & vbCrLf
    Response.Write "                             " & vbCrLf
    Response.Write "End function                 " & vbCrLf
End Sub


Sub WriteGetClientInfo()

    Response.Write "'==================================================================================================================	 " & vbCrLf
    Response.Write "Sub GetClientInfo()    " & vbCrLf
    Response.Write "                       " & vbCrLf
    Response.Write "    Dim iA,iB,iC,iD    " & vbCrLf
    Response.Write "    Dim strLang        " & vbCrLf
    Response.Write "                       " & vbCrLf
    Response.Write "    Call WebPathInfo(iA,iB,iC,iD)   " & vbCrLf
    Response.Write "                                    " & vbCrLf
    Response.Write "    strLang = UCase(iD)             " & vbCrLf
    Response.Write "                                    " & vbCrLf
    Response.Write "    frmLogin.txtLangCdForURL.value = strLang               " & vbCrLf
    Response.Write "    frmLogin.txtURL.value  = iA & ""//"" & iB & ""/"" & iC " & vbCrLf
    Response.Write "    frmLogin.txtVD.value   = iC                            " & vbCrLf
    Response.Write "    frmLogin.txtHttpWebSvrIPURL.value = iA & ""//"" & iB   " & vbCrLf
    Response.Write "	                                            " & vbCrLf
    Response.Write "    If strLang = ""TEMPLATE"" Then              " & vbCrLf
    Response.Write "       frmLogin.txtLang.value = ""KO""          " & vbCrLf
    Response.Write "    Else                                        " & vbCrLf
    Response.Write "       frmLogin.txtLang.value = strLang         " & vbCrLf
    Response.Write "    End If                                      " & vbCrLf
    Response.Write "	                                            " & vbCrLf
    Response.Write "    Call CalculateMath(iA,iB,iC,iD)             " & vbCrLf
    Response.Write "                                                " & vbCrLf
    Response.Write "    frmLogin.txtClientNum1000.value        = iA " & vbCrLf
    Response.Write "    frmLogin.txtClientNumDec.value         = iB " & vbCrLf
    Response.Write "    frmLogin.txtClientDateFormat.value     = iC " & vbCrLf
    Response.Write "    frmLogin.txtClientDateSeperator.value  = iD " & vbCrLf
    Response.Write "                                                " & vbCrLf
    Response.Write "End Sub                                         " & vbCrLf

End Sub


Sub WriteEraser()

    Response.Write "'==================================================================================================================  " & vbCrLf
    Response.Write "Sub Eraser()           " & vbCrLf
    Response.Write "    Dim iRet           " & vbCrLf
    Response.Write "    Dim objConn        " & vbCrLf
    Response.Write "    Dim iTest          " & vbCrLf

    Response.Write "    On Error Resume Next                                  " & vbCrLf
    Response.Write "    Set objConn = CreateObject(""uniConnector.cGlobal"")  " & vbCrLf
    
    Response.Write "    iTest = objConn.chkConn                               " & vbCrLf

    Response.Write "    If Trim(err.Description) = """" And iTest > 0 Then                                                                  " & vbCrLf
    Response.Write "        MsgBox Replace(""클라인언트 PC의 %1용 공통 파일을 제거하기 위해서는 사용 중인 %1를 모두 종료해야 합니다."",""%1"",""" & Request.Cookies("unierp")("gLogoName") & """) , vbInformation, iLogo   " & vbCrLf
    Response.Write "        Exit Sub                                                                                                        " & vbCrLf
    Response.Write "    End If                                                                                                              " & vbCrLf
    Response.Write "    iRet = MsgBox(Replace(""PC의 uniERPII 공통 CAB을 제거후 재설치 하시겠습니까?"",""%1"",""" & Request.Cookies("unierp")("gLogoName") & """) , vbYesNo       + vbQuestion, gLogo)     " & vbCrLf
    Response.Write "    If iRet = vbYes Then                         " & vbCrLf
    Response.Write "       top.location.href = ""uniRemoveCAB.asp?asp="" & frmLogin.initp.value  " & vbCrLf
    Response.Write "    End If        " & vbCrLf
    Response.Write "End Sub           " & vbCrLf

End Sub

Sub WriteGetSystemInfo()

    Response.Write "'==================================================================================================================  " & vbCrLf
    Response.Write "Function GetSystemInfo() " & vbCrLf
    Response.Write "	Dim MDAC             " & vbCrLf
    Response.Write "	                     " & vbCrLf
    Response.Write "	If MSIEVer < ""6.0"" Then " & vbCrLf
    Response.Write "	   MsgBox ""uniERPII를 사용하기위해서는 상위 Internet Explorer 버전을 필요로 합니다."" & VbCrLf  & VbCrLf & ""상위의 Internet Explorer 버전 설치를 시작 합니다."",vbExclamation,gLogo " & vbCrLf
    Response.Write "	   top.window.location.href = ""UpdateClient.asp?vbvbdd=A&adsad="" & ieVER " & vbCrLf
    Response.Write "	   Exit Function " & vbCrLf
    Response.Write "	End If " & vbCrLf
	
    Response.Write "	MDAC = ConnectorControl.GetDACVersion()  " & vbCrLf

    Response.Write "	If MDAC < ""2.8"" Then " & vbCrLf
    Response.Write "	   MsgBox ""uniERPII를 사용하기위해서는 상위 MDAC 버전을 필요로 합니다."" & VbCrLf  & VbCrLf & ""상위의 MDAC 버전 설치를 시작 합니다."",vbExclamation,gLogo " & vbCrLf
    Response.Write "	   top.window.location.href = ""UpdateClient.asp?vbvbdd=B&adsad="" & MDAC " & vbCrLf
    Response.Write "	   Exit Function " & vbCrLf
    Response.Write "	End If " & vbCrLf
	
    Response.Write "	Call GetClientInfo() " & vbCrLf

    Response.Write "End Function " & vbCrLf

End Sub

Sub WriteCheckIDPWD()

    Response.Write "'==================================================================================================================" & vbCrLf
    Response.Write "Function CheckIDPWD()" & vbCrLf

    Response.Write "    CheckIDPWD = False " & vbCrLf

    Response.Write "	If txtUsrIdview.value = """" Then                          " & vbCrLf
    Response.Write "       MsgBox ""아이디를 입력하세요!"",vbExclamation,gLogo " & vbCrLf
    Response.Write "       txtUsrIdview.focus                                      " & vbCrLf
    Response.Write "       Exit Function                                           " & vbCrLf
    Response.Write "	End If                                                     " & vbCrLf
	
    Response.Write "	If txtpwdview.value = """" Then                              " & vbCrLf
    Response.Write "       MsgBox ""비밀번호를 입력하세요!"",vbExclamation,gLogo " & vbCrLf
    Response.Write "       txtpwdview.focus                                          " & vbCrLf
    Response.Write "       Exit Function                                             " & vbCrLf
    Response.Write "	End If                                                       " & vbCrLf

    Response.Write "    frmLogin.txtUsrId.value = txtUsrIdview.value               " & vbCrLf
    Response.Write "    frmLogin.txtPWD.value   = MXD(txtpwdview.value)            " & vbCrLf   '2005-09-28
    Response.Write "    frmLogin.IDSave.value   = IDSaveview.checked               " & vbCrLf
    
    Response.Write "	If IDSaveview.checked Then                                     " & vbCrLf
    Response.Write "       frmLogin.ChkIdSave.value = True                         " & vbCrLf
    Response.Write "	Else                                                       " & vbCrLf
    Response.Write "       frmLogin.ChkIdSave.value = False                        " & vbCrLf
    Response.Write "	End If                                                     " & vbCrLf
	
    Response.Write "    CheckIDPWD = True	                                       " & vbCrLf
	
    Response.Write "End Function                                                   " & vbCrLf

End Sub


Sub WriteMultiCompany()

    Response.Write "'==================================================================================================================" & vbCrLf
    Response.Write "Sub LoadCompInfoD()" & vbCrLf

    Response.Write "    Dim iCompanyMajorList" & vbCrLf
    Response.Write "    Dim iCompanyMinorList" & vbCrLf
    Response.Write "    Dim iLoop            " & vbCrLf 

    Response.Write "	iCompanyMajorList = Split(""" & iCompanyAndDdList & """,Chr(12)) " & vbCrLf
	
    Response.Write "	ReDim gCompanyList(UBound(iCompanyMajorList) - 1,1)            " & vbCrLf
	
    Response.Write "	For iLoop = 0 To UBound(iCompanyMajorList) - 1                 " & vbCrLf
    Response.Write "	    iCompanyMinorList = Split(iCompanyMajorList(iLoop), "";"") " & vbCrLf
	    
    Response.Write "	    gCompanyList(iLoop,0) = iCompanyMinorList(0)   " & vbCrLf
    Response.Write "	    gCompanyList(iLoop,1) = iCompanyMinorList(1)   " & vbCrLf

    Response.Write "	Next   " & vbCrLf

    Response.Write "    iCompanyMajorList = gCompanyList(0,1)              " & vbCrLf

    Response.Write "	For iLoop = 0 To UBound(gCompanyList,1)            " & vbCrLf
    Response.Write "        Call SetCombo(CompanyList,gCompanyList(iLoop,0),gCompanyList(iLoop,0))   " & vbCrLf
    Response.Write "	Next                                                                         " & vbCrLf

    Response.Write "    Call SetDBList(CompanyList.value)   " & vbCrLf

    Response.Write "End Sub   " & vbCrLf

    Response.Write "Sub LoadCompInfoG()   " & vbCrLf
    Response.Write "    Dim iTemp         " & vbCrLf
    Response.Write "    Dim iLoop         " & vbCrLf

    Response.Write "	gCompanyList = Split(""" & iCompanyAndDdList & """,""::"") " & vbCrLf

        
    Response.Write "	For iLoop = 0 To UBound(gCompanyList)              " & vbCrLf
    Response.Write "	    If Trim(gCompanyList(iLoop)) <> """" Then      " & vbCrLf
    Response.Write "           iTemp = Split(gCompanyList(iLoop),""<>"")   " & vbCrLf
    Response.Write "	       If Trim(iTemp(1)) > """" Then               " & vbCrLf 
    Response.Write "              Call SetCombo(CompanyList,iTemp(0) & ""_"" & iTemp(1) ,iTemp(2))           " & vbCrLf
    Response.Write "           Else                                                                          " & vbCrLf 
    Response.Write "              Call SetCombo(CompanyList,iTemp(0) & ""_"" & iTemp(1) ,iTemp(0) & ""_NA"") " & vbCrLf
    Response.Write "           End If                                                                        " & vbCrLf
    Response.Write "        End If                                                                           " & vbCrLf
    Response.Write "	Next                                                                                 " & vbCrLf

    Response.Write "End Sub " & vbCrLf

    Response.Write "Sub CompanyList_OnChange()                " & vbCrLf
    Response.Write "	If gDebugMode = ""Y"" Then            " & vbCrLf
    Response.Write "       Call SetDBList(CompanyList.value)  " & vbCrLf
    Response.Write "    End If                                " & vbCrLf
    Response.Write "End Sub                                   " & vbCrLf  

    Response.Write "Sub SetDBList(ByVal pCompanyCD)    " & vbCrLf
    Response.Write "    Dim index         " & vbCrLf
    Response.Write "    Dim iDataExists   " & vbCrLf

    Response.Write "    Dim Temp    " & vbCrLf
    Response.Write "    Dim Temp2   " & vbCrLf
    Response.Write "    Dim iLoop   " & vbCrLf

    Response.Write "    index = FindCompanyIndex(pCompanyCD)   " & vbCrLf
    
    Response.Write "    Temp = gCompanyList(index,1)     " & vbCrLf

    Response.Write "	For iLoop = 0 To DbList.length   " & vbCrLf
    Response.Write "        DbList.remove(0)             " & vbCrLf
    Response.Write "	Next                             " & vbCrLf

    Response.Write "    Temp = Split(Temp , "":"")       " & vbCrLf
    
    Response.Write "    iDataExists = False              " & vbCrLf
    
    Response.Write "	For iLoop = 0 To UBound(Temp)      " & vbCrLf
    Response.Write "	    Temp2 = Trim(Temp(iLoop))      " & vbCrLf
    Response.Write "	    If Temp2 > """" Then           " & vbCrLf 
    Response.Write "	       Temp2 = Split(Temp2,""."")  " & vbCrLf
    Response.Write "           Call SetCombo(DBList,Temp2(1),Temp2(1)) " & vbCrLf
    Response.Write "           iDataExists = True                      " & vbCrLf
    Response.Write "        End If                                     " & vbCrLf
    Response.Write "	Next                                           " & vbCrLf

    Response.Write "    If iDataExists = False Then                    " & vbCrLf
    Response.Write "        Call SetCombo(DBList,""NONE"",""N/A"")     " & vbCrLf
    Response.Write "    End If                                         " & vbCrLf
    
    Response.Write "End Sub                                            " & vbCrLf

    Response.Write "Function FindCompanyIndex(ByVal pCompanyCD)        " & vbCrLf
    Response.Write "    Dim iLoop                                      " & vbCrLf

    Response.Write "	For iLoop = 0 To UBound(gCompanyList,1)        " & vbCrLf
    Response.Write "        If gCompanyList(iLoop,0) = pCompanyCD Then " & vbCrLf
    Response.Write "           FindCompanyIndex = iLoop                " & vbCrLf
    Response.Write "           Exit Function                           " & vbCrLf
    Response.Write "        End If                                     " & vbCrLf
    Response.Write "	Next                                           " & vbCrLf
    
    Response.Write "End Function                                       " & vbCrLf


    Response.Write "Sub SetCombo(pCombo, ByVal strValue, ByVal strText) " & vbCrLf
    Response.Write "	Dim objEl                                       " & vbCrLf
			
    Response.Write "	Set objEl = Document.CreateElement(""OPTION"")  " & vbCrLf	
    Response.Write "	objEl.Text = strText                            " & vbCrLf
    Response.Write "	objEl.Value = strValue                          " & vbCrLf

    Response.Write "	pcombo.Add(objEl)                               " & vbCrLf
    Response.Write "	Set objEl = Nothing                             " & vbCrLf

    Response.Write "End Sub                                             " & vbCrLf


End Sub


%>