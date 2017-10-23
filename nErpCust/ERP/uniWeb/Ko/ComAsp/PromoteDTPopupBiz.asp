<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<!-- #Include file="../inc/incSvrDate.inc" -->
<!-- #Include file="../inc/incSvrNumber.inc" -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOOCOOKIE","PB")                                                                      '☜: Clear Error status

    Dim txtStrDT 
    Dim iLoop,jLoop
    Dim arrField
    Dim TmpStr
    Dim strSQL,strData,strWhere
	Dim adoRec
	Dim isOverFlowKey
	Dim isOverFlowName
	Dim strDataTemp
	Dim tmp2By2Array
	
    Const C_SHEETMAXROWS = 30									'한화면에 보일수 있는 최대 Row 수 

    Call HideStatusWnd

    On Error Resume Next

    Err.Clear
    


    txtStrDT  = uniConvDate(Request("txtDate"))
    
    strWhere = Trim(Request("txtWhere"))
    
   
    strData = ""	
	
	if request("txtTable") = "1" then
	
		strSQL = " Select Top " & C_SHEETMAXROWS + 1 & " Resent_promote_dt, count(emp_no) " 
		strSQL = strSQL & " From  haa010t "
	
		if request("txtlgdate") <> "" then
			strSQL = strSQL & " where Resent_promote_dt <= '" & uniConvDate(request("txtlgdate")) &"' "
		else
			strSQL = strSQL & " Where Resent_promote_dt <= '" & txtStrDt &"' "
		end if
		strSQL = strSQL & " Group by Resent_promote_dt "
		strSQL = strSQL & " order by  Resent_promote_dt DESC"			
	
	else  
		strSQL = " Select Top " & C_SHEETMAXROWS + 1 & " promote_dt, count(emp_no) " 
		strSQL = strSQL & " From  hba080t "
	
		if request("txtlgdate") <> "" then
			strSQL = strSQL & " where promote_dt <= '" & uniConvDate(request("txtlgdate")) &"' "
		else
			strSQL = strSQL & " Where promote_dt <= '" & txtStrDt &"' "
		end if
		strSQL = strSQL & " Group by promote_dt "
		strSQL = strSQL & " order by  promote_dt DESC"
	end if
	
	
	Set adoRec = Server.CreateObject("ADODB.RecordSet")    
                                         ' adOpenForwardOnly, adLockReadOnly, adCmdTable
	adoRec.Open strSQL,gADODBConnString, 0                , 1             , 1	
	
    If Err.Number = 0 Then		
       If Not( adoRec.EOF And adoRec.BOF ) Then
          tmp2By2Array = adoRec.GetRows()
		  
          adoRec.Close 
          Set adoRec = Nothing
       
          For iLoop = 0 To UBound(tmp2By2Array,2)
              If iLoop < C_SHEETMAXROWS Then
                  For jLoop = 0 To UBound(tmp2By2Array,1) 
                      strDataTemp = tmp2By2Array(jLoop,iLoop)
                      if jLoop = 0 then
					     strDataTemp = uniConvDateAtoB(strDataTemp, gAPdateFormat, gDateformat)
					  end if
                      strData = strData & Chr(11) & strDataTemp  
                  Next 
                  strData = strData & Chr(11) & lgLngMaxRow + iLoop + 1    
                  strData = strData & Chr(11) & Chr(12)
              Else
                  isOverFlowKey  = tmp2By2Array(0,iLoop)
                  isOverFlowName = tmp2By2Array(1,iLoop)
                  Exit For
              End If
          Next
       End If   
    End If   
	
%>		    



<Script Language="vbscript">   

  On Error Resume Next  
	With parent
        .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"
        .lgDate        = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgCount        = "<%=ConvSPChars(isOverFlowName)%>"
        .vspdData.focus	        
        .DbQueryOk()
	End With

</Script>