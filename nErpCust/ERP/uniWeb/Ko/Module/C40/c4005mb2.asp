<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
<!--'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4005MA1.asp
'*  4. Program Name			:배부요소DATA등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4005MB1.asp
'*						
'*  7. Modified date(First)	: 2005/09/05
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************-->
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%

	Call LoadBasisGlobalInf()	
	Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     
     Call SubBizQuery()
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim sCostCd,sDstbCd,sGubun
	Dim tmpC1,arrCon1, arrCon2,iColCon,iRowCon
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt
	
	sStartDt	= Request("txtYYYYMM")		
	sNextKey	= Request("lgStrPrevKey")	
	sCostCd=  request("txtCode") 
	sDstbCd=  request("txtFctrCd")	
	sGubun =  request("txtGubun")  

	If sCostCd = "" Then sCostCd = "%"
	If sDstbCd = "" Then sDstbCd = "%"	


    With lgObjComm
		.CommandTimeout = 0
	
		.CommandText = "dbo.usp_C_C4005MA1_list"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DSTB_FCTR_CD",	adVarXChar,	adParamInput, 3,Replace(sDstbCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OPT_FLAG",	adVarXChar,	adParamInput, 2,Replace(sGubun, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
		    
        Set oRs = lgObjComm.Execute
        
    End With
       
    'Response.Write "Err=" & Err.Description
    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
	If sGubun="C" then 
		 If Not oRs is nothing Then
				
				arrRows = oRs.GetRows()
				iLngRowCnt = UBound(arrRows,2) 
				iLngColCnt	= UBound(arrRows, 1) 

				If iLngRowCnt ="" Then 		
					Response.Write " <Script Language=vbscript>	                        " & vbCr
					Response.Write " With parent                                        " & vbCr
					Response.Write "	.lgStrPrevKey2 = """"" & vbCr 	
					Response.Write " End With                                        " & vbCr
					Response.Write  " </Script>                  " & vbCr
					Exit Sub
				End If
				tmpC1=""
				sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
				For iLngRow = 0 To 	iLngRowCnt	
							  
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))					
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
						If arrRows(2, iLngRow)="%1" then
						iStrData = iStrData & Chr(11) & REPLACE(ConvSPChars(arrRows(2, iLngRow)),"%1","사내")
						Else
						iStrData = iStrData & Chr(11) & REPLACE(ConvSPChars(arrRows(2, iLngRow)),"%2","외주가공")
						end if
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
						iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(5, iLngRow),ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)
							

						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(6, iLngRow))				
						iStrData = iStrData & Chr(11) & Chr(12)
							
			Next
						
					
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr 	
			Response.Write " .frm1.hCode.value    = """ & ConvSPChars(Request("txtCode"))   & """" & vbCr
			Response.Write " .frm1.hFctrCd.value     = """ & ConvSPChars(Request("txtFctrCd"))    & """" & vbCr
			Response.Write " .frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
			Response.Write " .frm1.hYYYYMM.value     = """ & ConvSPChars(Request("txtYYYYMM"))      & """" & vbCr
			Response.Write "	.DbQueryOk " & sRowSeq+1-100 & ","""""  & vbCr 
			
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
		End If
	Else
		If Not oRs is nothing Then
				arrRows = oRs.GetRows()
				iLngRowCnt = UBound(arrRows,2) 
				iLngColCnt	= UBound(arrRows, 1) 
				
				If iLngRowCnt < 0 Then 		
					Response.Write " <Script Language=vbscript>	                        " & vbCr
					Response.Write " With parent                                        " & vbCr
					Response.Write "	.lgStrPrevKey = """"" & vbCr 	
					Response.Write " End With                                        " & vbCr
					Response.Write  " </Script>                  " & vbCr
					Exit Sub
				End If
				tmpC1=""
				sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
				For iLngRow = 0 To 	iLngRowCnt	
						If arrRows(2, iLngRow)="%1" then
						iStrData = iStrData & Chr(11) & REPLACE(ConvSPChars(arrRows(2, iLngRow)),"%1","사내")
						Else
						iStrData = iStrData & Chr(11) & REPLACE(ConvSPChars(arrRows(2, iLngRow)),"%2","외주가공")
						end if  
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))					
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))
						iStrData = iStrData & Chr(11) & ""
						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
						iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(5, iLngRow),ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)
							

						iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(6, iLngRow))				
						iStrData = iStrData & Chr(11) & Chr(12)
							
			Next

						
					
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 			
			Response.Write " .frm1.hCode.value    = """ & ConvSPChars(Request("txtCode"))   & """" & vbCr
			Response.Write " .frm1.hFctrCd.value     = """ & ConvSPChars(Request("txtFctrCd"))    & """" & vbCr
			Response.Write " .frm1.hGubun.value    = """ & ConvSPChars(Request("txtGubun")) & """" & vbCr
			Response.Write " .frm1.hYYYYMM.value     = """ & ConvSPChars(Request("txtYYYYMM"))      & """" & vbCr
			Response.Write "	.DbQueryOk " & sRowSeq+1-100 & ","""""   & vbCr 
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
		End If
 End IF
 				

       
End Sub	


%>

