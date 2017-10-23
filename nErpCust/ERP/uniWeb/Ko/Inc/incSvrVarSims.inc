<%   
   Select Case UCase(gLang)
      Case "KO","TEMPLATE"
               Response.CharSet = "euc-kr"                               'Korea
               gLogoName = "대사우서비스"
               gLogo = "ESS"
      Case "CN"
                Response.CharSet = "GB2312"                               'China
                gLogoName = "ESS"
                gLogo = "ESS"
      Case "JA"
                Response.CharSet = "shift_jis"                            'Japan
                gLogoName = "ESS"
                gLogo = "ESS"                
      Case "EN"
               'Response.CharSet = "windows-1252"                         'U.S.A
                gLogoName = "ESS"
                gLogo = "ESS"                
   End Select       
     
%>
