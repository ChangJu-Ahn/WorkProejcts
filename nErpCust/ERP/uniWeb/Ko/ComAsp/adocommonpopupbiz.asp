<%Option Explicit%>
<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
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
    Dim isCondition
    Dim isCondition1
    DIM isWhere
    
    Const C_SHEETMAXROWS = 30									'한화면에 보일수 있는 최대 Row 수 
    
    Call HideStatusWnd
    
    If Trim(ggAmtOfMoney.DecPoint) = "" Then
       Call LoadInfTB19029B("Q","*","NOCOOKIE","MB")
    End If

    strWhichSide = 1

    txtNextCode  = Trim(Request("txtNextCode"))
    txtNextName  = Trim(Request("txtNextName"))
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
	
	isWhere				= Request("isWhere")
	isCondition1		= Request.Form("txtisCondition")
	'isWhere = 1 ' 0은 >= 조건 1은 like 조건 
	if Request.QueryString("isFlag") = 1 Then isWhere = 0 End If
    
    strSQL = ""
    
    For iLoop = 0 To intDataCount - 1
        strSQL = strSQL & arrField(iLoop) & ","
    Next

    strSQL = Left(strSQL,Len(Trim(strSQL)) - 1)
    arrStrDT  = Split(Request("arrStrDT"),gColSep)    
	
	If gDBKind = "ORACLE" Then
       strSQL =          " Select Distinct " & strSQL
    Else
       strSQL =          " Select Distinct Top " & C_SHEETMAXROWS + 1 & " " & strSQL
    End If

	strSQL = strSQL & " From   " & strTable
	strSQL = strSQL & " Where  " & strWhere	
	
	If strWhere <> "" Then
	   strSQL = strSQL & " And " 
	End If
    
	If iNameDuplication = "F" Then

        If txtNextCode <> "" Then
        
			IF isWhere = 0 THEN
				strSQL = strSQL & arrField(0) & ">= " & FilterVar(txtNextCode, "''", "S") & " order by " & arrField(0)		
			ELSE
				if isCondition1 = "" then
				strSQL = strSQL & arrField(0) & " LIKE " & FilterVar("%" & txtNextCode & "%", "''", "S") & " order by " & arrField(0)		
				else
				strSQL = strSQL & arrField(0) & " LIKE " & FilterVar("%" & txtNextCode & "%", "''", "S") & " AND " & arrField(0) & " NOT IN (" & isCondition1 & ")  order by " & arrField(0)		
				end if
            END IF
        ElseIf txtNextName <> "" Then
            strWhichSide = 2
            IF isWhere = 0 THEN
	            strSQL = strSQL & arrField(1) & ">= " & FilterVar(txtNextName, "''", "S") & " order by " & arrField(1) & ", " & arrField(0)
            ELSE
				if isCondition1 = "" then
		        strSQL = strSQL & arrField(1) & " LIKE " & FilterVar("%" & txtNextName & "%", "''", "S") & " order by " & arrField(1) & ", " & arrField(0)
		        ELSE
		        strSQL = strSQL & arrField(1) & " LIKE " & FilterVar("%" & txtNextName & "%", "''", "S") & " AND " & arrField(0) & " NOT IN (" & isCondition1 & ")  order by " & arrField(0)		
		        END IF
            END IF
        Else   
			IF isWhere = 0 THEN
				strSQL = strSQL & arrField(0) & ">= " & FilterVar(txtNextCode, "''", "S") & " order by " & arrField(0)		
			ELSE
				if isCondition1 = "" then
				strSQL = strSQL & arrField(0) & " LIKE " & FilterVar("%" & txtNextCode & "%", "''", "S") & " order by " & arrField(0)		
				ELSE
				strSQL = strSQL & arrField(0) & " LIKE " & FilterVar("%" & txtNextCode & "%", "''", "S") & " AND " & arrField(0) & " NOT IN (" & isCondition1 & ")  order by " & arrField(0)		
				END IF
            END IF
        End If   
        
    Else
        strWhichSide = 2
        strSQL = strSQL & "(( " & arrField(1) & "= " & FilterVar(txtNextName, "''", "S") & " and " & arrField(0) & ">= " & FilterVar(txtNextCode, "''", "S") & " ) or " & _
                   arrField(1) & "> " & FilterVar(txtNextName, "''", "S") & " ) " & " order by " & arrField(1) & ", " & arrField(0)
    End If
    
	If gDBKind = "ORACLE" Then
       strSQL = "Select * from ( " & Replace(strSQL,"''","' '") & " ) a "
       strSQL = strSQL & "Where rownum <= " & C_SHEETMAXROWS + 1
    End If




	Set adoRec = Server.CreateObject("ADODB.RecordSet")    
                                         ' adOpenForwardOnly, adLockReadOnly, adCmdTable
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
                      Select Case arrStrDT(jLoop)
                         Case "DD"  :    strDataTemp = UNIDateClientFormat(strDataTemp)
                         Case "F2"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggAmtOfMoney.DecPoint, 0)
                         Case "F3"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggQty.DecPoint       , 0)
                         Case "F4"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggUnitCost.DecPoint  , 0)
                         Case "F5"  :    strDataTemp = UNINumClientFormat (strDataTemp, ggExchRate.DecPoint  , 0)
                      End Select 
                    
                      strData = strData & Chr(11) & strDataTemp  
                  Next    
                  
				  IF isWhere = 1 THEN 'LIKE 검색인 경우만 적용됨 
					if txtNextCode <> "" or txtNextName <> "" then
						IF iLoop = C_SHEETMAXROWS - 1 THEN
						isCondition = isCondition & "'" & CSTR(tmp2By2Array(0,iLoop)) & "'"
						ELSE
						isCondition = isCondition & "'" & tmp2By2Array(0,iLoop) & "',"
						END IF
					end if
                  END IF
                                    
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
 '   Call WriteToLog(strSQL)

  IF isCondition1 = "" THEN
  ELSE
	isCondition = isCondition1 & "," & isCondition
  END IF
  
  
Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    pPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLang & "\Log"

    pPath = pPath & "\ADOCommonPopup" & "[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gUsrID & "][" & Request.ServerVariables("REMOTE_ADDR") & "].txt"

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
  
  parent.document.all("txtisCondition").value   = "<%=isCondition%>"
  
	With parent
        .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"
        .lgStrNextCodeKey      = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgStrNextNameKey      = "<%=ConvSPChars(isOverFlowName)%>"
        .lgNameDuplication   = "<%=iNameDuplication%>"        
        .DbQueryOk()
	End With
	

</Script>