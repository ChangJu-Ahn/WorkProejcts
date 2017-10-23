<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<!-- #Include file="../inc/incSvrDate.inc" -->
<!-- #Include file="../inc/incSvrNumber.inc" -->
<%
    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOOCOOKIE","PB")                                                                      '☜: Clear Error status

    Dim arrStrDT 
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
    
    ReDim arrField(8)
	
	lgLngMaxRow       = Request("txtMaxRows")   
    arrStrDT  = Split(Request("arrStrDT"),gColSep)   
    strWhere = Trim(Request("txtWhere"))
    
    arrField(0) = Trim(Request("arrField1"))
    arrField(1) = Trim(Request("arrField2"))
    arrField(2) = Trim(Request("arrField3"))
    arrField(3) = Trim(Request("arrField4"))
    arrField(4) = Trim(Request("arrField5"))
    arrField(5) = Trim(Request("arrField6"))
    arrField(6) = Trim(Request("arrField7"))
	arrField(7) = Trim(Request("arrField8"))              
	arrField(8) = Trim(Request("arrField9"))	'cyc
    strSQL = ""

    For iLoop = 0 To UBound(arrField)
        If arrField(iLoop) <> "" Then
           If InStr(1,UCase(arrField(iLoop)),"CONVERT") Then
              strSQL = strSQL &  Mid( arrField(iLoop),1,Len( arrField(iLoop)) - 1)    & ",21),"
           Else
              strSQL = strSQL & arrField(iLoop) & ","
           End If   
        End If   
    Next

    strSQL = Mid(strSQL,1,Len(strSQL) - 1)
	
	strSQL =          " Select Top " & C_SHEETMAXROWS + 1 & " " & strSQL
	strSQL = strSQL & " From   " & Request("txtTable")
	strSQL = strSQL & " Where  " & strWhere	
	
	If strWhere <> "" Then
	   strSQL = strSQL & " And " 
	End If
	
	If Request("txtCd") <> "" Then 		
		strSQL = strSQL & arrField(0) & ">= " & FilterVar(Request("txtCode"),"''","S")  & " order by " & arrField(0)				
    ElseIf Request("txtNm") <>"" then    		
		strSQL = strSQL & arrField(1) & ">= " & FilterVar(Request("txtName"),"''","S")  & " order by " & arrField(1)		
	Else		
		strSQL = strSQL & arrField(0) & ">= " & FilterVar(Request("txtCode"),"''","S")  & " order by " & arrField(0)				
	End if

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
                  For jLoop = 0 To UBound(tmp2By2Array,1) - 1
                      strDataTemp = tmp2By2Array(jLoop,iLoop)
                 
                      If InStr(1, UCase(arrField(jLoop)), "CONVERT") > 0 And  InStr(1, UCase(arrField(jLoop)), "CHAR") > 0 Then  ' If numeric or date
                         If Instr(1, strDataTemp, "-") > 0 Then
                            strDataTemp = UniConvDateAToB(Mid(strDataTemp,1,10),gServerDateFormat,gDateFormat)
                         End If   
                      End If

                      Select Case arrStrDT(jLoop)
                         Case "DD"  :    strDataTemp = UNIDateClientFormat(strDataTemp)
                         Case "F2"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggAmtOfMoney.DecPoint, 0)
                         Case "F3"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggQty.DecPoint       , 0)
                         Case "F4"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggUnitCost.DecPoint  , 0)
                         Case "F5"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggExchRate.DecPoint  , 0)
                      End Select 
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
        .lgCode        = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgName        = "<%=ConvSPChars(isOverFlowName)%>"
        .vspdData.focus	        
        .DbQueryOk()
	End With

</Script>