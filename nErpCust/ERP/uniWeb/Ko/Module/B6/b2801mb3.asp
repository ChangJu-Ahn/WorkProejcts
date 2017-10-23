<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : b2801mb3.asp	
'*  4. Program Name         : Storage Location Delete
'*  5. Program Desc         :
'*  6. Comproxy List        : +B28011ManageStorageLocation

'*  7. Modified date(First) : 2000/04/25
'*  8. Modified date(Last)  : 2000/04/25
'*  9. Modifier (First)     : Kweon Soon Ho
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'*                            this mark(¢Á) Means that "may  change"
'*							  this mark(¡Ù) Means that "must change"
'* 13. History              : -1999/09/12 : ..........
'*                            
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%															
Call LoadBasisGlobalInf()

On Error Resume Next

Err.Clear	

Call HideStatusWnd											

Dim pPB6G010	

Dim I3_b_storage_location
Dim iCommandSent

Const I001_I3_sl_cd = 0

ReDim I3_b_storage_location(I001_I3_sl_cd)			

Dim StrNextKey										
Dim lgStrPrevKey									
Dim LngRow
Dim GroupCount 


    If Request("txtSLCd1") = "" Then
		Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
	


	'-----------------------
	'Data manipulate area
	'-----------------------									
	I3_b_storage_location(I001_I3_sl_cd) = Request("txtSLCd1")
	iCommandSent = "DELETE"
	
	Set pPB6G010 = Server.CreateObject("PB6G010.cBManageStorageLoc")
	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call pPB6G010.B_MANAGE_STORAGE_LOCATION(gStrGlobalCollection, iCommandSent, , , I3_b_storage_location)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set pPB6G010 = Nothing										
		Response.End
	End If

	
	Set pPB6G010 = Nothing



%>
<Script Language=vbscript>
	With parent															
		.DbDeleteOk
	End With
</Script>
