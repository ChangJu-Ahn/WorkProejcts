<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : 角荤急喊包府  历厘 诀公 贸府 ASP
'*  2. Function Name        : 
'*  3. Program ID           : i2131mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PI2G050 I_COUNT_PHY_INV

'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2002/074/05
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(⑿) means that "Do not change"
'*                            this mark(⒘) Means that "may  change"
'*                            this mark(≠) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%	
	Call LoadBasisGlobalInf()											
    
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd 

	Dim pPI2G050															
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i
	
	Dim I1_i_physical_inventory_header_phy_inv_no
    Dim E1_i_physical_inventory_header_phy_inv_no    

	
	'-----------------------
	'Data manipulate area
	'-----------------------										

    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")

	I1_i_physical_inventory_header_phy_inv_no	= UCase(Request("txtCondPhyInvNo"))

	Set pPI2G050 = Server.CreateObject("PI2G050.cICountPhyInv")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
		Response.End
	End If
			
	Call pPI2G050.I_COUNT_PHY_INV(gStrGlobalCollection, _
								I1_i_physical_inventory_header_phy_inv_no, _
								itxtSpread)
	
	If CheckSYSTEMError(Err, True) = True Then
		Set pPI2G050 = Nothing														
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
		Response.End
	End If
			
	Set pPI2G050 = Nothing														
	
	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
  	Response.Write "    .RemovedivTextArea	"	& vbCr
  	Response.Write "    .DbSaveOk			"	& vbCr
	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End 			

%>
