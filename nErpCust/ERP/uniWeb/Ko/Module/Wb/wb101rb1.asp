<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	Dim isOverFlowKey,isOverFlowName
	Dim iNameDuplication
	Dim iTempLastName
    Dim strSQL,strTable,strWhere
    Dim arrStrDT 
	Dim tmp2By2Array
    Dim iLoop,jLoop
	Dim strDataTemp
	Dim strData
	Dim adoRec
    Dim arrField(6)
    Dim strWhichSide
    Dim txtNextCode,txtNextName
    Dim intDataCount
	
    Const C_SHEETMAXROWS = 30									'한화면에 보일수 있는 최대 Row 수 

    Call HideStatusWnd
    
    If Trim(ggAmtOfMoney.DecPoint) = "" Then
       Call LoadInfTB19029B("Q","*","NOCOOKIE","MB")
    End If
    
    strWhichSide = 1
    
    txtNextCode  = Replace(Trim(Request("txtNextCode")),"'","''")
    txtNextName  = Replace(Trim(Request("txtNextName")),"'","''")
    iNameDuplication = Request("NameDuplication")
    
    strTable     = Request("txtTable")
    strWhere     = Request("txtWhere")
    
    arrField(0) = Trim(Request("arrField1"))
    arrField(1) = Trim(Request("arrField2"))
    arrField(2) = Trim(Request("arrField3"))
    arrField(3) = Trim(Request("arrField4"))
    arrField(4) = Trim(Request("arrField5"))
    arrField(5) = Trim(Request("arrField6"))
    arrField(6) = Trim(Request("arrField7"))
              
	intDataCount = Request("gintDataCnt")

    strSQL = ""

    For iLoop = 0 To intDataCount - 1 
        If arrField(iLoop) <> "" Then
           If InStr(1,UCase(arrField(iLoop)),"CONVERT") Then
              strSQL = strSQL &  Mid( arrField(iLoop),1,Len( arrField(iLoop)) - 1)    & ",21),"
           Else
              strSQL = strSQL & arrField(iLoop) & ","
           End If   
        End If   
    Next

    strSQL = Left(strSQL,Len(Trim(strSQL)) - 1)
    arrStrDT  = Split(Request("arrStrDT"),gColSep)   
    
	If gDBKind = "ORACLE" Then
       strSQL =          " Select     " & strSQL
	Else
	   strSQL =          " Select DISTINCT Top " & C_SHEETMAXROWS + 1 & " " & strSQL
	End If
	
	strSQL = strSQL & " From   " & strTable
	strSQL = strSQL & " Where  " & strWhere	
	
	If strWhere <> "" Then
	   strSQL = strSQL & " And " 
	End If
	
	If iNameDuplication = "F" Then
        If txtNextCode <> "" Then
            strSQL = strSQL & arrField(0) & ">= '" & txtNextCode & "'" & " order by " & arrField(0)		
        ElseIf txtNextName <> "" Then
            strWhichSide = 2
            strSQL = strSQL & arrField(1) & ">= '" & txtNextName & "'" & " order by " & arrField(1) & ", " & arrField(0)
        Else   
            strSQL = strSQL & arrField(0) & ">= '" & txtNextCode & "'" & " order by " & arrField(0)		
        End If   
    Else
        strWhichSide = 2
        strSQL = strSQL & "(( " & arrField(1) & "= '" & txtNextName & "' and " & arrField(0) & ">= '" & txtNextCode & "' ) or " & _
                   arrField(1) & "> '" & txtNextName & "' ) " & " order by " & arrField(1) & ", " & arrField(0)
    End If

	If gDBKind = "ORACLE" Then
       strSQL =  "select  *  from ( " & Replace(strSQL,"''","' '") & " ) a "
       strSQL =   strSQL & "Where rownum <= " & C_SHEETMAXROWS + 1
    End If
    
	Set adoRec = Server.CreateObject("ADODB.RecordSet")    
                                         ' adOpenForwardOnly, adLockReadOnly, adCmdTable
	On Error Resume Next
	Response.Write strSQL
	adoRec.Open strSQL,gADODBConnString, 0                , 1             , 1	

    If Err.Number = 0 Then		
       If Not( adoRec.EOF And adoRec.BOF ) Then
          isOverFlowKey  = ""
          isOverFlowName = ""
          strData        = ""
          iTempLastName = ""
          iNameDuplication = "F"
       
          tmp2By2Array = adoRec.GetRows()
       
          adoRec.Close 
          Set adoRec = Nothing
       
          For iLoop = 0 To UBound(tmp2By2Array,2)
              If iLoop < C_SHEETMAXROWS Then
                  For jLoop = 0 To UBound(tmp2By2Array,1) 
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
                  strData = strData & Chr(11) & Chr(12)
                  If iLoop = C_SHEETMAXROWS-1 And strWhichSide = 2 Then
                      iTempLastName = tmp2By2Array(1,iLoop)
                  End If
              Else              
                  If strWhichSide = 1 Then 
                     isOverFlowKey  = tmp2By2Array(0,iLoop)
                  Else   
                     isOverFlowName = tmp2By2Array(1,iLoop)
                     If isOverFlowName = iTempLastName Then
						iNameDuplication = "T"
						isOverFlowKey  = tmp2By2Array(0,iLoop)                     
                     End If
                  End If
                     
                  Exit For
              End If
          Next
       End If   
    End If   

 '   Call WriteToLog("Next Key [" & isOverFlowKey & "][" & isOverFlowName & "]")
  '  Call WriteToLog(strSQL)

Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    pPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLang & "\Log"

    pPath = pPath & "\CommonPopup" & "[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gUsrID & "][" & Request.ServerVariables("REMOTE_ADDR") & "].txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile( pPath,8,True)
       
    objFile.WriteLine pLogData
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

End Sub
	
%>		    

<Script Language="vbscript">   

  On Error Resume Next  
	With parent
        .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"
        .lgStrNextCodeKey   = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgStrNextNameKey   = "<%=ConvSPChars(isOverFlowName)%>"
        .lgNameDuplication   = "<%=iNameDuplication%>"
        .DbQueryOk()
	End With

</Script>