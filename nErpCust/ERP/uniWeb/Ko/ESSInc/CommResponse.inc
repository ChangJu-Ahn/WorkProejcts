<% 
   Dim URLLangUserID
   Dim iUrlTempCommResponse
   Dim tLogo

   Const NodeNm1     = "LoadBasisGlobalInf"
   Const NodeNm2     = "GetGlobalInf"
   Const NodeNm3     = "Login"
   Const gCharSet    = "D"  'U : unicode, D:DBCS
   Const gCharSQLSet = "D"  'U : unicode, D:DBCS
   
   Const gMSXMLDOMDocument  = "Msxml2.DOMDocument.4.0"
   
   URLLangUserID = Request.Cookies("unierp")("gURLLangUserID")
   
   Response.Cookies("unierp")("gCharSet")    = gCharSet
   Response.Cookies("unierp")("gCharSQLSet") = gCharSQLSet
     
   if URLLangUserID <> "" then 
	  URLLangUserID = " " & URLLangUserID
   end if 

   Select Case UCase(Request.Cookies("unierp")("gLang"))
      Case "KO","TEMPLATE"
                Response.CharSet = "euc-kr"                               'Korea
                
                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "uniERPII"
                   Response.Cookies("unierp")("gLogo")     = "uniERPII" & URLLangUserID 
                End If   
                
                Response.Cookies("unierp")("gFontName") = "±¼¸²Ã¼" 
                Response.Cookies("unierp")("gFontSize") = "9" 
      Case "CN"
                Response.CharSet = "GB2312"                               'China

                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "Bizentro"
                   Response.Cookies("unierp")("gLogo")     = "Bizentro" & URLLangUserID 
                End If  

                Response.Cookies("unierp")("gFontName") = "Arial" 
                Response.Cookies("unierp")("gFontSize") = "9" 
      Case "JA"
                Response.CharSet = "shift_jis"                            'Japan

                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "Bizentro"
                   Response.Cookies("unierp")("gLogo")     = "Bizentro" & URLLangUserID 
                End If  

                Response.Cookies("unierp")("gFontName") = "Arial" 
                Response.Cookies("unierp")("gFontSize") = "9" 
      Case "IR"
                Response.CharSet = "windows-1256"                            'Japan

                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "Bizentro"
                   Response.Cookies("unierp")("gLogo")     = "Bizentro" & URLLangUserID 
                End If  

                Response.Cookies("unierp")("gFontName") = "Traffic" 
                Response.Cookies("unierp")("gFontSize") = "9" 
      Case "EN"
            '    Response.CharSet = "windows-1252"                         'U.S.A
               
                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "Bizentro"
                   Response.Cookies("unierp")("gLogo")     = "Bizentro" & URLLangUserID 
                End If  
                
                Response.Cookies("unierp")("gFontName") = "Arial" 
                Response.Cookies("unierp")("gFontSize") = "9" 
      Case "HU"
                Response.CharSet = "windows-1250"                         'Hungary

                tLogo = Request.Cookies("unierp")("gLogoName")
                
                If tLogo <> "" Then
                   Response.Cookies("unierp")("gLogoName") = tLogo
                   Response.Cookies("unierp")("gLogo")     = tLogo & URLLangUserID 
                Else 
                   Response.Cookies("unierp")("gLogoName") = "Bizentro"
                   Response.Cookies("unierp")("gLogo")     = "Bizentro" & URLLangUserID 
                End If  

                Response.Cookies("unierp")("gFontName") = "Arial" 
                Response.Cookies("unierp")("gFontSize") = "9" 
   End Select               

   If gCharSet = "U" Then
      Session.CodePage = 65001
      Response.CharSet = "utf-8"
   End If
   
   iUrlTempCommResponse= request.servervariables("path_info")  
   iUrlTempCommResponse = Split(iUrlTempCommResponse,"/")
   Response.Cookies("unierp").path = "/" & iUrlTempCommResponse(1) & "/" & iUrlTempCommResponse(2)  

 
   
%>
