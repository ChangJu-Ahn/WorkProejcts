<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2312rb1.asp
'*  4. Program Name         : 재고현황조회 팝업 처리 
'*  5. Program Desc         : 
'*  6. Modified date(First) : Park , Bumsoo
'*  7. Modified date(Last)  : Ahn , Jungje
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd

On Error Resume Next

Dim ADF          
Dim strRetMsg        
Dim UNISqlId, UNIValue, UNILock, UNIFlag 
Dim rs0, rs1, rs2, rs3
Dim PvArr
Dim iCnt
Dim iRCnt     

Dim strData
Dim lgMaxCount
Dim lgStrPrevKey

Dim strItemCd
Dim StrSlCd

Err.Clear                
	
	lgMaxCount		= 100
	lgStrPrevKey	= Request("lgStrPrevKey")

	'=======================================================================================================
	' 만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saf"
	UNISqlId(2) = "180000sad"
	 
	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(1, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtSlCd"), "''", "S")

	UNILock = DISCONNREAD : UNIFlag = "1"
	 
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	' Plant 명 Display
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing

		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write "	parent.txtPlantNm.value = """" " & vbCr  	   	  
		Response.Write "</Script>      " & vbCr   
		Response.End   

	Else
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write "	parent.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """ " & vbCr  	   	  
		Response.Write "</Script>      " & vbCr   
	
		rs1.Close
		Set rs1 = Nothing
	End If

	' 품목명 Display
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing

		Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write "	parent.txtItemNm.value = """" " & vbCr  	   	  
		Response.Write "	parent.txtItemCd.focus " & vbCr  	   	  
		Response.Write "</Script>      " & vbCr   
		Response.End   
	Else
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write "	parent.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCr  	   	  
		Response.Write "	parent.txtItemSpec.value = """ & ConvSPChars(rs2("SPEC")) & """" & vbCr  	   	  
		Response.Write "	parent.txtSafetyStock.value = """ & UniConvNumberDBToCompany(rs2("SS_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & """" & vbCr  	   	  
		Response.Write "</Script>      " & vbCr   

		rs2.Close
		Set rs2 = Nothing
	End If

	' 창고명 Display
	IF Request("txtSlCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing

			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript> " & vbCr   
			Response.Write "	parent.txtSLNm.value = """" " & vbCr  	   	  
			Response.Write "	parent.txtSLCd.focus " & vbCr  	   	  
			Response.Write "</Script>      " & vbCr   
			Response.End   
		Else
			Response.Write "<Script Language=vbscript> " & vbCr   
			Response.Write "	parent.txtSLNm.value = """ & ConvSPChars(rs3("SL_NM")) & """" & vbCr  	   	  
			Response.Write "</Script>      " & vbCr   

			rs3.Close
			Set rs3 = Nothing
		End If
	End IF


	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "I2312RB1"
	 
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtSlCd") = "" Then
		strSlCd = "|"
	Else
		StrSlCd = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strSlCd

	UNILock = DISCONNREAD : UNIFlag = "1"

	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End            
	End If
	
	iCnt	= 0
	strData	= ""
	
	If Len(Trim(lgStrPrevKey)) Then                                      
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If
	
	ReDim PvArr(0)
	
	For iRCnt = 1 to iCnt * lgMaxCount
		rs0.MoveNext
	Next
	
	iRCnt = -1
	
	Do while Not (rs0.EOF Or rs0.BOF)
		iRCnt = iRCnt + 1	
	
		ReDim Preserve PvArr(iRCnt)
		
		strData	=	Chr(11) & ConvSPChars(rs0("sl_cd")) & _
					Chr(11) & ConvSPChars(rs0("sl_nm")) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("good_on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("schd_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("schd_issue_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("avail_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("stk_in_trns_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("allocation_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & CStr((iCnt * lgMaxCount) + iRCnt) & Chr(11) & Chr(12)
					
		PvArr(iRCnt) = strData
		
		If  iRCnt >= lgMaxCount Then
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	strData = Join(PvArr, "")
	
	If  iRCnt < lgMaxCount Then                                          
        lgStrPrevKey = ""                                                
    End If
					
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing           

    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr

	Response.Write "	.ggoSpread.Source	= .vspdData1 "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr
	
	Response.Write "	.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr  	   	  
  	Response.Write "	.hItemCd.value  = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCr
  	Response.Write "	.hSlCd.value    = """ & ConvSPChars(Request("txtSlCd")) & """" & vbCr
	Response.Write "	.lgStrPrevKey   = """ & ConvSPChars(lgStrPrevKey)	   & """" & vbCr  
	
  	Response.Write "	If .vspdData1.MaxRows < .parent.VisibleRowCnt(.vspdData1, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

	
	Response.Write "End with	" & vbCr
    Response.Write "</Script>      " & vbCr   

	Response.End   
%>


