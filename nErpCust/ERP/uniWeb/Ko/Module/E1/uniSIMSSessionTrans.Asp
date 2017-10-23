
<%

   Response.Cookies("uniSIMS")("gUsrId")      = Request("uid")
   Response.Cookies("uniSIMS")("gPwd")        = Request("pwd")  
   Response.Cookies("uniSIMS")("gLang")       = Request("lang")  
   Response.Cookies("uniSIMS")("gVD")         = Request("VD")
   
   Response.Cookies("unierp")("gClientDateFormat")      = Request("ClientDateFormat")    
   Response.Cookies("unierp")("gClientDateSeperator")   = Request("ClientDateSeperator")    
   Response.Cookies("unierp")("gClientNum1000")         = Request("ClientNum1000")    
   Response.Cookies("unierp")("gClientNumDec")			= Request("ClientNumDec")      
   Response.Cookies("unierp")("gXMLFileNm")				= Request("strXMLFileNm")   			
   Response.Cookies("unierp")("gHTTPXMLFileNm")			= Request("strHTTPXMLFileNm") 
  
   Response.Redirect "elogin_ok.asp"

%>
