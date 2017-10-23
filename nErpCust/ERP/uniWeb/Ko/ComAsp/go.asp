<%
   Dim strGo
   Dim strASPName
   Dim strFolder
   Dim strASPNameNm

   On Error Resume Next

   Err.Clear 

   strGo = Trim(Request("txtGo"))
                  ' 0                1                2         3  
   strSQL =          " Select a.mnu_id, a.upper_mnu_id,  b.mnu_nm, c.called_frm_id "
   strSQL = strSQL & " From z_auth_gen a, z_lang_co_mast_mnu b , z_co_mast_mnu c   "
   strSQL = strSQL & " Where a.mnu_id   = b.mnu_id "
   strSQL = strSQL & " And   a.mnu_id   = c.mnu_id "
   strSQL = strSQL & " And   a.mnu_id   = '" & strGo                              & "'"
   strSQL = strSQL & " And   b.lang_cd  = '" & Request.Cookies("unierp")("gLang") & "'"


   Set pRs = Server.CreateObject("ADODB.Recordset")   

   pRs.Open strSQL,Request.Cookies("unierp")("gADODBConnString"),0,1

   If Err.number = 0 Then
      If Not (pRs.EOF Or pRs.BOF) Then
         strASPName   = pRs(3)
         strFolder    = pRs(1)
         strASPNameNm = pRs(2)
         pRs.Close
         Set pRs = Nothing
         Response.Redirect "../Module/" & strFolder & "/" & strASPName & ".asp?strRequestMenuID=" & UCase(strGo)  & "&strASPMnuMnuNm="& strASPNameNm
      End If   
   End If   
%>
