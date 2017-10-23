<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Distribution Factor
'*  3. Program ID           :
'*  4. Program Name         : 오더배부규칙 정보 등록 
'*  5. Program Desc         : 오더배부규칙 정보 등록  
'*  6. Modified date(First) : 2005/10/24
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

DIM lgCopyVersion
	lgCopyVersion= request("versionFlag")
	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     IF 	lgCopyVersion="Y" THEN
         Call SubBizQuery()
     ELSE
		Call SubBizQueryMulti()
     END if
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim iSTrCode, iSTrCode2, IntRetCD
	Dim lgCopyVersion,strMsg_cd
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	iSTrCode= Trim(ucase(Request("txtVER_CD")))
	iSTrCode2= Trim(ucase(Request("hVER_CD")))
		
	With lgObjComm
			.CommandTimeout = 0
			
			.CommandText = "dbo.usp_C4004MA1_CopyVersion"		' --  변경해야할 SP 명 
		    .CommandType = adCmdStoredProc

			.Parameters.Append lgObjComm.CreateParameter("RETURN_VAL",  adInteger,adParamReturnValue)	' -- No 수정 

			' -- 변경해야할 조회조건 파라메타들 
			.Parameters.Append lgObjComm.CreateParameter("@NEW_CD",	adVarXChar,	adParamInput, 3,Replace(iSTrCode, "'", "''"))
			.Parameters.Append lgObjComm.CreateParameter("@COPY_CD",	adVarXChar,	adParamInput, 3,Replace(iSTrCode2, "'", "''"))
			.Parameters.Append lgObjComm.CreateParameter("@INSRT_USER_ID"			,adVarXChar,adParamInput,13, gUsrId)
			.Parameters.Append lgObjComm.CreateParameter("@UPDATE_USER_ID"			,adVarXChar,adParamInput,13, gUsrId)
			.Parameters.Append lgObjComm.CreateParameter("@MSG_CD",	adVarXChar,	adParamOutput, 6)
			
		    .Execute  ,, adExecuteNoRecords
		    
		End With

		If Instr( Err.Description , "B_MESSAGE") > 0 Then
			strMsg_cd = lgObjComm.Parameters("@msg_cd").Value  
			If HandleBMessageError(vbObjectError, Err.Description,"" , "") = True Then
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr		
				Response.Write "parent.fncNew " & vbCr    
				Response.Write "</Script>"
				Exit Sub
			End If
		Else
			If CheckSYSTEMError(Err, True) = True Then	
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr		
				Response.Write "parent.fncNew " & vbCr    
				Response.Write "</Script>"

				Exit Sub
			End If
		End If
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.lgCopyVersion="""& lgCopyVersion & """"  & vbcr
		Response.Write "	.frm1.versionFlag.value="""& lgCopyVersion & """"  & vbcr
		Response.Write "	Call .FncQuery " & vbCr 	
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr	

End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
   Dim iStrData, iStrData2
 
   'Dim iLngRow,iLngCol
    
    Dim iStrCode, sNextKey
    Dim lgCopyVersion,i
    
	Dim oRs, sTxt, arrRows, iLngRow, iLngCol,   sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim arrCol, iColDept, iiColSize
	
    Const c_i1_ver_cd = 0
	Const c_i1_flag =1
	Const c_i1_flag_nm =2
	Const c_i1_wc_cd = 3
	Const c_i1_wc_nm = 4
	Const c_i1_group_level = 5
	Const c_i1_acct_group = 6
	Const c_i1_acct_group_nm = 7
	Const c_i1_acct_cd = 8
	Const c_i1_acct_nm = 9
	Const c_i1_adstb_fctr = 10
	Const c_i1_adstb_fctr_nm = 11
	Const c_i1_sdstb_fctr = 12
	Const c_i1_sdstb_fctr_nm = 13
	Const c_i1_row_seq = 14
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Const C_SHEETMAXROWS_D  = 100    

	sNextKey = Trim(Request("lgStrPrevKey"))	                                                         '☜: Clear Error status

	iStrCode= Trim(ucase(Request("txtVER_CD")))

	With lgObjComm
			.CommandTimeout = 0
			
			.CommandText = "dbo.usp_C4004MA1_LIST"		' --  변경해야할 SP 명 
		    .CommandType = adCmdStoredProc

			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

			' -- 변경해야할 조회조건 파라메타들 
			.Parameters.Append lgObjComm.CreateParameter("@VER_CD",	adVarXChar,	adParamInput, 3,Replace(iStrCode, "'", "''"))			
			.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, C_SHEETMAXROWS_D)	
			.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
			.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 

        Set oRs = lgObjComm.Execute
	End With

    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If
	
	If oRs.Eof and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
		

    If Not oRs.EOF Then
		arrRows = oRs.GetRows()

		iLngRowCnt = UBound(arrRows, 2) 
		iLngColCnt	= UBound(arrRows, 1) 
	
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)

		For iLngRow = 0 To 	iLngRowCnt
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_ver_cd,iLngRow))		   		  		                                                
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_flag,iLngRow))  		
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_flag_nm,iLngRow))  	
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_wc_cd,iLngRow))  		
				iStrData = iStrData & chr(11) & ""        
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_wc_nm,iLngRow))		
				If trim(arrRows(c_i1_group_level,iLngRow)) =0 Then
				istrData = istrData & Chr(11) & " "
				Else
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_group_level,iLngRow))	
				End If
				iStrData = iStrData & chr(11) & ""        				
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_acct_group,iLngRow))   			
				iStrData = iStrData & chr(11) & ""
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_acct_group_nm,iLngRow))     		
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_acct_cd,iLngRow))   		
				iStrData = iStrData & chr(11) & ""		
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_acct_nm,iLngRow))    
				
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_adstb_fctr,iLngRow))       
				iStrData = iStrData & chr(11) & ""
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_adstb_fctr_nm,iLngRow))   
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_sdstb_fctr,iLngRow))       
				iStrData = iStrData & chr(11) & ""
				istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_sdstb_fctr_nm,iLngRow))   				         

				iStrData = iStrData & Chr(11) & arrRows(c_i1_row_seq, iLngRow)	' -- Row_Seq					
				iStrData = iStrData & gRowSep				
				
		Next
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
		
		If sNextKey <> "*" and iLngRowCnt >=99  Then
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 	
		else		
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
		End If
		
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr		
		
    ElseIf sNextKey = "" Then 
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If    
End Sub   
%>

