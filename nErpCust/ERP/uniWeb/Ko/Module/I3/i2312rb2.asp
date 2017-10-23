<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2312rb2.asp
'*  4. Program Name         : 재고현황조회 팝업 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
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
Dim rs0          
Dim PvArr
Dim iCnt
Dim iRCnt     

Dim strData
Dim lgMaxCount
Dim lgStrPrevKey2

Dim strItemCd
Dim strSlCd
	lgMaxCount		= 100
	lgStrPrevKey2	= Request("lgStrPrevKey2")

Err.Clear              

	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "I2312RB2"
	 
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
		rs0.Close
		Set rs0 = Nothing
		Response.End             
	End If
	
	iCnt	= 0
	strData	= ""
	
	If Len(Trim(lgStrPrevKey2)) Then                                     
       If Isnumeric(lgStrPrevKey2) Then
          iCnt = CInt(lgStrPrevKey2)
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
		
		strData =	Chr(11) & ConvSPChars(rs0("tracking_no")) & _
					Chr(11) & ConvSPChars(rs0("lot_no")) & _
					Chr(11) & ConvSPChars(rs0("lot_sub_no")) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("good_on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("stk_on_trns_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(rs0("picking_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & ConvSPChars(rs0("block_indicator")) & _
					Chr(11) & CStr((iCnt * lgMaxCount) + iRCnt) & Chr(11) & Chr(12)
					
		PvArr(iRCnt) = strData
		
		If  iRCnt >= lgMaxCount Then
            iCnt = iCnt + 1
            lgStrPrevKey2 = CStr(iCnt)
            Exit Do
        End If	
		rs0.MoveNext
	Loop
	
	strData = Join(PvArr, "")
	
	If  iRCnt < lgMaxCount Then                                          
        lgStrPrevKey2 = ""                                                
    End If

	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing           

    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr

	Response.Write "	.ggoSpread.Source	= .vspdData2 "										& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"							& vbCr
	
	Response.Write "	.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """"		& vbCr  	   	  
  	Response.Write "	.hItemCd.value  = """ & ConvSPChars(Request("txtItemCd")) & """"		& vbCr
  	Response.Write "	.hSlCd.value    = """ & ConvSPChars(Request("txtSlCd")) & """"			& vbCr
	Response.Write "	.lgStrPrevKey   = """ & ConvSPChars(lgStrPrevKey2)	   & """" & vbCr  
	
  	Response.Write "	If .vspdData2.MaxRows < .parent.VisibleRowCnt(.vspdData2, 0)And .lgStrPrevKey2 <> """" Then	"	& vbCr
  	Response.Write "		.DbDtlQuery(.vspdData1.ActiveRow)								"	& vbCr
  	Response.Write "    Else								"									& vbCr
  	Response.Write "		.DbDtlQueryOk()								"						& vbCr
	Response.Write "    End If								"									& vbCr

	
	Response.Write "End with	"																& vbCr
    Response.Write "</Script>      "															& vbCr   

	Response.End   

%>

