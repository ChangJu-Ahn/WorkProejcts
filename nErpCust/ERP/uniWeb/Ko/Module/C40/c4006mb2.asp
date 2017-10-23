<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%''**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4006MA1.asp
'*  4. Program Name			:완성품환산율등록 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4006Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/08/29
'*  8. Modified date(Last)	: 2005/11/03
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: HJO
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)

	Call SubBizQueryMulti()
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
   Dim iStrData
  
    Dim sYYYYMM, sNextKey,sWcCd,sOrderNo,sPlant
    
	Dim oRs,  arrRows, iLngRow, iLngCol,   sRowSeq, iLngRowCnt, iLngColCnt
	
	Const c_i1_wc_cd = 0
	Const c_i1_wc_nm =1
	Const c_i1_item_cd = 2
	Const c_i1_item_nm = 3
	Const c_i1_order_no = 4
	Const c_i1_prod_rate = 5
	Const C_I1_ROW_SEQ = 6
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Const C_SHEETMAXROWS_D  = 100    

	sNextKey = Trim(Request("lgStrPrevKey"))	                                                         '☜: Clear Error status

	sYYYYMM=replace(request("txtYYYYMM"),"-","")
	sWcCd=request("txtWcCd")
	sOrderNo=  request("txtProdOrderNo") 
	sPlant =  request("txtPlantCd")  
	
	If len(sWcCd)=0 then sWcCd="%"	
	If len(sOrderNo)=0 then sOrderNo="%"

	With lgObjComm
			.CommandTimeout = 0
			
			.CommandText = "dbo.usp_C4006MA1_LIST"		' --  변경해야할 SP 명 
		    .CommandType = adCmdStoredProc

			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

			' -- 변경해야할 조회조건 파라메타들 
			.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sYYYYMM, "'", "''"))
			.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 7,Replace(sWcCd, "'", "''"))			
			.Parameters.Append lgObjComm.CreateParameter("@ORDER_NO",	adVarXChar,	adParamInput, 18,Replace(sOrderNo, "'", "''"))			
			.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput,4,Replace(sPlant, "'", "''"))						
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
		

    If Not oRs.EOF Then
		arrRows = oRs.GetRows()

		iLngRowCnt = UBound(arrRows, 2) 
		iLngColCnt	= UBound(arrRows, 1) 
	

		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
		istrData=""
		For iLngRow = 0 To 	iLngRowCnt
			istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_wc_cd,iLngRow ))		   
			iStrData = iStrData & chr(11) & ""                                                                      
			istrData = istrData & Chr(11) & ConvSPChars(arrRows(c_i1_wc_nm,iLngRow))  
			istrData = istrData & Chr(11) & ConvSPChars(arrRows( c_i1_Item_cd,iLngRow))		   
			iStrData = iStrData & chr(11) & ""                                                                      
			istrData = istrData & Chr(11) & ConvSPChars(arrRows( c_i1_item_nm,iLngRow))  
			istrData = istrData & Chr(11) & ConvSPChars(arrRows( c_i1_order_no,iLngRow))       
			iStrData = iStrData & chr(11) & ""
			istrData = istrData & Chr(11) & UNINumClientFormat(arrRows(c_i1_prod_rate,iLngRow),ggQty.DecPoint,0) * 100
      
 			iStrData = iStrData & Chr(11) & arrRows(C_I1_ROW_SEQ, iLngRow)	' -- Row_Seq					
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
		elseif sNextKey <> "*" then
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
		End If
		Response.Write " .frm1.hPlantCd.value    = """ & ConvSPChars(Request("txtPlantCd"))   & """" & vbCr
		Response.Write " .frm1.hWcCd.value     = """ & ConvSPChars(Request("txtWcCd"))    & """" & vbCr
		Response.Write " .frm1.hProdOrderNo.value    = """ & ConvSPChars(Request("txtProdOrderNo")) & """" & vbCr
		Response.Write " .frm1.hYYYYMM.value     = """ & ConvSPChars(Request("txtYYYYMM"))      & """" & vbCr
		'Response.Write " .DbQueryOk " & intARows & ",iMaxRow"   & vbCr 
		
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
		
		
    ElseIf sNextKey = "" Then 
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If    
End Sub   
%>

