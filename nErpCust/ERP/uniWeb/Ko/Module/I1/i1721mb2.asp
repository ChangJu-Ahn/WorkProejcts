<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Batch Posting ��� asp
'*  2. Function Name        : 
'*  3. Program ID           : i1721mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : i17111BatchPostGoodsMovsvr

'*  7. Modified date(First) : 2001/05/14
'*  8. Modified date(Last)  : 2001/05/14
'*  9. Modifier (First)     : Lee Hae Ryong
'* 10. Modifier (Last)      : Lee Hae Ryong
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%	
	Call LoadBasisGlobalInf()

	On Error Resume Next	
	    
    Call HideStatusWnd
    
	Dim iPI1G200			
	Dim LngRow, iMaxRow
	
	Dim arrRowVal, arrColVal			

	Dim I1_ief_supplied_select_char
	Dim IG1_import_group
		Const I135_IG1_I1_item_document_no	= 0
		Const I135_IG1_I1_document_year		= 1
	Dim iErrorPosition
    ReDim iErrorPosition(0)
    Dim iErrorPositionIndex

 	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i

    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    
    itxtSpreadArrCount = -1
             
	ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")
    
	'-----------------------
	'Data manipulate area
	'-----------------------					
	
    I1_ief_supplied_select_char = "D"
	  
	If itxtSpread <> "" Then
	
		arrRowVal = Split(itxtSpread, gRowSep)
		iMaxRow = UBound(arrRowVal) -1
		
		ReDim IG1_import_group(iMaxRow, I135_IG1_I1_document_year)
		ReDim iErrorPositionIndex(iMaxRow)
		
		For LngRow = 0 To iMaxRow
		    		
			arrColVal = Split(arrRowVal(LngRow), gColSep)
		
          	IG1_import_group(LngRow, I135_IG1_I1_item_document_no) = arrColVal(0)
			IG1_import_group(LngRow, I135_IG1_I1_document_year)	   = arrColVal(1)
			iErrorPositionIndex(LngRow)                            = arrColVal(2)
				
		Next
		
		Set iPI1G200 = Server.CreateObject("PI1G200.cIBchPostGoodMovSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err,True) = True Then
		      Set iPI1G200 = Nothing			
			  Response.End						
		End If

		Call iPI1G200.I_BATCH_POST_GOODS_MOV_SVR(gStrGlobalCollection, _
											     I1_ief_supplied_select_char, _
											     IG1_import_group, _
											     iErrorPosition)

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err,True) = True Then
			Set iPI1G200 = Nothing			
      		If iErrorPosition(0) <> 0 Then
				Response.Write "<Script Language=VBScript>" & vbCrLF
				Response.Write "Call parent.SheetFocus(" & iErrorPositionIndex(iErrorPosition(0)) & ", 1)" & vbCrLF
				Response.Write "</Script>" & vbCrLF
			End If
			Response.End
		End If
		                   
		Set iPI1G200 = Nothing	
				
	End If

%>

<Script Language=vbscript>
	With parent	
		.DbSaveOk
	End With
</Script>
