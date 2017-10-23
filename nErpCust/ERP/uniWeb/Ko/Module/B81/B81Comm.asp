<%Response.Buffer = True  %>

<%
'=========================================================='
'Recordset을 배열로 받아 그리드에 뿌리기 
'==========================================================
Sub ListupDataGrid(pgetArr,dataFormatCol)
	Dim strData
	Dim i,j
		for i=0 to uBound(pgetArr,2)
			for j=0 to uBound(pgetArr,1)
			
			if inStr(dataFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvDateDbToCompany(pgetArr(j,i),"")
			else
				strData = strData & Chr(11) & ConvSPChars(pgetArr(j,i))
			end if	
			
		
			next 
			strData =  strData & Chr(11) & i &  Chr(11) & Chr(12) 
		next 
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
		Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr   
		Response.Write "	.ggoSpread.SSShowData     """ & strData	 & """" & ",""F""" & vbCr
		Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr 
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr   
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
End Sub	



  										                                               
  Function GetNameChk1(fld,tbl,where,pValue,obj)
	dim tSql
    if len(trim(pValue))<1 then exit function 'If data exists then exit
    on error resume next 
		tSql="select "&fld&" from "&tbl&" where "& where
		
		If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") <> False Then       'If data exists	
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "."&obj&".value=""" &lgObjRs(0)&"""" & vbCr
			Response.Write "End with" & vbCr
			Response.Write "</Script>"		& vbCr
			lgObjRs.close
 
		else 
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "."&obj&".value=""""" & vbCr
			Response.Write "End with" & vbCr
			Response.Write "</Script>"		& vbCr
		End If
	
  End Function 

'=========================================================================================
' Name : GetNameChk
' Desc : main 에 codeName set하기 , display msg
'=========================================================================================

 Function GetNameChk(fld,tbl,where,pValue,obj,msg,msgYN)
  'on error resume next 
	dim tSql
	
    if len(trim(pValue))=0 then
		call goObjClear(obj)
		exit function 'If data exists then exit
    end if 
   
     
		tSql="select "&fld&" from "&tbl&" where "& where
	
		'-------------------------------------
		'display message
		'-------------------------------------
		if msgYN="Y" then
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") <> False Then       'If data exists	
			    Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "With parent.frm1" & vbCr
				Response.Write "."&obj&"_nm.value=""" &lgObjRs(0)&"""" & vbCr
				Response.Write "End with" & vbCr
				Response.Write "</Script>"		& vbCr
				lgObjRs.close
			else 
			     Call DisplayMsgBox("970000", vbInformation, Msg, "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			     call goObjClear(obj)
			     call goFocus(obj)
			     lgObjRs.close
			     Response.End 
				exit Function
				

				
			End If
		else 
		'-------------------------------------
		'only set nameValue
		'-------------------------------------
			If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") <> False Then       'If data exists	
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "With parent.frm1" & vbCr
				Response.Write "."&obj&"_nm.value=""" &lgObjRs(0)&"""" & vbCr
				Response.Write "End with" & vbCr
				Response.Write "</Script>"		& vbCr
				lgObjRs.close
				
 
			else 
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "With parent.frm1" & vbCr
				Response.Write "."&obj&"_nm.value=""""" & vbCr
				Response.Write "End with" & vbCr
				Response.Write "</Script>"		& vbCr
			End If
	
	
		end if
		
  End Function 
  
  
  
'=========================================================================================
' Name : GetNameChkGrid
' Desc : main 에 codeName set하기 , display msg
'=========================================================================================

 Function GetNameChkGrid(fld,tbl,where,row,col,objGrid,msg)
	dim tSql
  
  '  on error resume next 
  
		tSql="select "&fld&" from "&tbl&" where "& where

			If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") <> False Then       'If data exists	
			     lgObjRs.close
			else 
			     Call DisplayMsgBox("970000", vbInformation, msg, "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			     call goFocusGRid(objGrid,row,col)
			     lgObjRs.close
			     Response.End 
				exit Function
				

				
			End If
		
  End Function 
'=========================================================================================
' Name : fnCheckItem
' Desc : 
'=========================================================================================

 Sub fnCheckItem(rs,objItem,msg)
    
		If  rs.EOF And rs.BOF Then
		    rs.Close
		    Set rs1 = Nothing
		  
			   Call DisplayMsgBox("970000", vbInformation, Msg, "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			   call goObjClear(objItem)
			   call goFocus(objItem)
			   Response.End 
				exit sub
			
		Else  
			call goSetName( objItem & "_nm",rs(1))  
		    rs.Close
		    Set rs = Nothing
		End If
    end Sub
    
'=========================================================================================
' Name : goSetName
' Desc : 
'=========================================================================================
Sub goSetName(obj,sVal)

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1."&obj&".value="""&sVal&""" " & vbCr
	Response.Write "</Script>" & vbCr
			
end Sub

'=========================================================================================

'=========================================================================================

Sub goFocus(str)

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1."&str&".focus" & vbCr
	Response.Write "</Script>" & vbCr
			
end Sub

Sub goObjClear(str)

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1."&str&"_nm.value="""" " & vbCr
	Response.Write "</Script>" & vbCr
			
end Sub

Sub goFocusGrid(objGrid,row,col)

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write objGrid&".row=" & row & vbCr
	Response.Write objGrid&".col=" & col & vbCr
	Response.Write objGrid&".action=0" & vbCr
	Response.Write "</Script>" & vbCr
			
end Sub







Function CheckSystemErrorY(objError, pBool,pTitle)

    Dim iDesc

    CheckSystemErrorY = False
    
    If objError.Number = 0 Then
       Exit Function
    End If
    
    CheckSystemErrorY = True
    
    If objError.Number = vbObjectError Then
      If InStr(UCase(objError.Description), "B_MESSAGE") > 0 Then
         If HandleBMessageError(objError.Number, objError.Description, pTitle, "") = True Then
            Exit Function
         End If
      End If
    End If

    objError.Clear
    
End Function










  
		%>
