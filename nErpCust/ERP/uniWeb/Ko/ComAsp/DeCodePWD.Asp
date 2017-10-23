<%

   Dim iRet
   
   On Error Resume Next
   Err.Clear

   Select Case UCase(Request("FLAG"))
      Case "ESS", ""
			iRet = Request("PWD")
   
			Set Enc2 = Server.CreateObject("EDCodeCom.EDCodeObj.1")
   
			If Err.number <> 0 Then
			   Response.Write "X"
			Else
			   Response.Write Enc2.Decode(iRet)
			   
			End If
			
			Set Enc2 =  Nothing		
			
	  Case "KWAVE"
			iRet = Request("PWD")

			Set Enc = server.CreateObject("kwave_code.decode")
   
			If Err.number <> 0 Then
			   Response.Write "X"
			Else
			   Response.Write Enc.DecodeData(iRet)			   
			End If
						
			Set Enc =  Nothing		
			
	End Select
	
%>