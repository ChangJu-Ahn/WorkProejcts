<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : Single Sample
*  3. Program ID           : processsvr
*  4. Program Name         : processsvr
*  5. Program Desc         : processsvr
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : Kim Hwa Young
* 11. Comment              :
=======================================================================================================-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/incSvrMain.asp"  -->
<!-- #Include file="../inc/incSvrDate.inc"  -->
<!-- #Include file="../inc/incSvrNumber.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/ProcessStyle.css">

<%
Call HideStatusWnd

On Error Resume Next
Err.Clear

'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()        
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

%>
<SCRIPT LANGUAGE=VBSCRIPT>

	With parent
		.DbQueryOk
	End With

</Script>

	<TABLE ALIGN=Center>
		<TR ALIGN=Center>
			<TD>
				<TABLE>
					<TR ALIGN=Center>
<%
							Call QueryProcess()
%>
					</TR>
				</TABLE>
			</TD>
		</TR>
	</TABLE>

<%

Sub QueryProcess()

	On Error Resume Next
	Err.Clear
	
	Const strTagTD1 = "<TD CLASS=""ProcessTD"" WIDTH=* ALIGN=Center><A href=vbscript:Parent.PgmJump("""
	Const strTagTD12 = "<TD CLASS=""ProcessTD2"" WIDTH=* ALIGN=Center><A href=vbscript:Parent.PgmJump("""
	Const strTagTD2 = """) title="""
	Const strTagTD21 = """>"
	Const strTagTD3 = "</TD>"
		
	Const strTagTR = "</TR></TABLE></TD></TR><TR ALIGN=Center><TD ALIGN=Center><TABLE><TR ALIGN=Center><TD ALIGN=Center><img src=""../../CShared/image/process/c5.gif""></TD></TR></TABLE></TD></TR><TR ALIGN=Center><TD><TABLE><TR ALIGN=Center><TD>"

	
	
	Dim strPrcsSeq
	Dim strPrcsSubSeq	
	Dim strMnuID
	Dim strMnuNm
	Dim strOptnFlag
	Dim strRemark
	
	Dim strPrevSeq	
	Dim strTag	
	
	Dim istrCode
	Dim E1_Z_Prcs_Mnu_Asso
	Dim iZC0014
	
	Const C_SHEETMAXROWS_D = 100
	Const ZC14_E1_PRCS_SEQ     = 0
	Const ZC14_E1_PRCS_SUB_SEQ = 1
	Const ZC14_E1_MNU_ID       = 2
	Const ZC14_E1_MNU_NM       = 3
	Const ZC14_E1_OPTN_FLAG    = 4
	Const ZC14_E1_REMARK       = 5
	Const ZC14_E1_PRCS_CD      = 6
	Const ZC14_E1_PRCS_NM      = 7
	Const ZC14_E1_SYS_FLAG     = 8
	
	istrCode = ConvSPChars(FilterVar(Request("txtPrcsCd"),"","SNM"))
	
    Set iZC0014 = Server.CreateObject("PZCG014.cListPrcsMnuAsso")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iZC0014 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub													'☜: 비지니스 로직 처리를 종료함 
    End If

    E1_Z_Prcs_Mnu_Asso = iZC0014.ZC_PRCS_MNU_ASSO (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
	
    
	If CheckSYSTEMError(Err,True) = True Then		
       Set iZC0014 = Nothing	        
       Exit Sub
    End If
       	
	Set iZC0014 = Nothing
		
	For iLngRow =0 To UBound(E1_Z_Prcs_Mnu_Asso,2)
		
			strPrcsSeq = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_PRCS_SEQ,iLngRow))
			strPrcsSubSeq = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_PRCS_SUB_SEQ,iLngRow))			
			strMnuID = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_MNU_ID,iLngRow))
			strMnuNm = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_MNU_NM,iLngRow))
			strOptnFlag = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_OPTN_FLAG,iLngRow))
			strRemark = ConvSPChars(E1_Z_Prcs_Mnu_Asso(ZC14_E1_REMARK,iLngRow))
			
			
			'If strPrcsSeq <> "010" And strPrevSeq <> strPrcsSeq Then   'khy20030111
			If strPrevSeq <> "" And strPrevSeq <> strPrcsSeq Then
				strTag = strTag + strTagTR
			End If
			
			If strOptnFlag = "1" Then
				strTag = strTag + strTagTD12 + Trim(strMnuID) + strTagTD2 + strRemark + strTagTD21 + strMnuNm + strTagTD3				
			Else
				strTag = strTag + strTagTD1 + Trim(strMnuID) + strTagTD2 + strRemark + strTagTD21 + strMnuNm + strTagTD3
			End If

			strPrevSeq = strPrcsSeq			

	Next
	    
	
    
    Response.Write strtag
    
End Sub

%>
