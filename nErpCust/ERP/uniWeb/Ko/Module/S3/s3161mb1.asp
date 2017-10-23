<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3161MB1
'*  4. Program Name         : ����Ҵ� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*							  
'*  7. Modified date(First) : 2002/11/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho inkuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :     
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%													

On Error Resume Next														

Call HideStatusWnd

Dim iLngRow	
Dim iLngMaxRow
Dim istrData
Dim iStrPrevKey
Dim iStrNextKey

Dim arrValue
Dim istrMode

Dim ObjPS3G162
Const C_SHEETMAXROWS_D  = 30	

Dim imp_next_inv_alloc
Redim imp_next_inv_alloc(2)

Const C_imp_next_so_no = 0
Const C_imp_next_so_seq = 1
Const C_imp_next_schd_no = 2 

Dim C_inv_alloc
Redim C_inv_alloc(8)

Dim EG1_exp_grp
Const EG1_AvaInv			= 0			'�������Ȯ�� 
Const EG1_PurReqAuto		= 1			'���ſ�û�ڵ����� 

Dim EG2_exp_grp													
Const EG2_ItemCd			= 0			'ǰ��		
Const EG2_ItemName			= 1			'ǰ��� 
Const EG2_ItemSpec			= 2			'�԰� 
Const EG2_TrackingNo		= 3			'Tracking No
Const EG2_SoUnit			= 4			'���� 
Const EG2_SoQty				= 5			'���ַ� 
Const EG2_PreAllocQty		= 6			'���Ҵ緮 
Const EG2_BonusQty			= 7			'������	
Const EG2_PreAllocBonusQty	= 8			'���Ҵ������	
Const EG2_PromiseDt			= 9		'������� 
Const EG2_DlvyDt			= 10		'������	
Const EG2_SlCd				= 11		'â���ڵ� 
Const EG2_SlNm				= 12		'â��� 
Const EG2_PlantCd			= 13		'�����ڵ� 
Const EG2_PlantNm			= 14		'����� 
Const EG2_SoNo				= 15		'���ֹ�ȣ 
Const EG2_SoSeq				= 16		'���ּ��� 
Const EG2_SchdNo			= 17		'��ǰ���� 
Const EG2_PrePurReqQty		= 18		'�ⱸ�ſ�û��(Hidden)
Const EG2_GiQty				= 19		'������ 

istrMode = Request("txtMode")												' ���� ���¸� ���� 

Select Case istrMode

Case CStr(UID_M0001)														' ���� ��ȸ/Prev/Next ��û�� ���� 

    Err.Clear     
    
    '--------------------------------------------------------------------------------------------------------
    ' ����Ҵ�	������ �о�´�.
    '--------------------------------------------------------------------------------------------------------

    C_inv_alloc(0)	= Trim(Request("txtFromConSoNo"))    
    C_inv_alloc(1)	= Trim(Request("txtToConSoNo"))  
    C_inv_alloc(2)	= Trim(Request("txtShipToParty"))  
    C_inv_alloc(3)	= Trim(Request("txtSalesGrp"))  
    C_inv_alloc(4)	= Trim(Request("txtItem"))  
    C_inv_alloc(5)	= Trim(Request("txtPlant"))  
    C_inv_alloc(6)	= UNIConvDate(Trim(Request("txtFromDate")))
    C_inv_alloc(7)	= UNIConvDate(Trim(Request("txtToDate")))  
    C_inv_alloc(8)	= Trim(Request("txtRadio"))  
	
    iStrPrevKey = Trim(Request("lgStrPrevKey"))  
    
    If iStrPrevKey <> "" then
    
		arrValue = Split(iStrPrevKey, gColSep)
				
		imp_next_inv_alloc(C_imp_next_so_no) = Trim(arrValue(0))
		imp_next_inv_alloc(C_imp_next_so_seq) = Trim(arrValue(1))
		imp_next_inv_alloc(C_imp_next_schd_no) = Trim(arrValue(2))
		
    Else    
		imp_next_inv_alloc(C_imp_next_so_no) = ""
		imp_next_inv_alloc(C_imp_next_so_seq) = 0
		imp_next_inv_alloc(C_imp_next_schd_no) = 0
    End If    
	
    Set ObjPS3G162 = Server.CreateObject("PS3G162.CsListSchdForDlvySvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If

    Call ObjPS3G162.S_LIST_SCHD_FOR_DLVY_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, imp_next_inv_alloc, _
											 C_inv_alloc, EG1_exp_grp, EG2_exp_grp)
    
   
    '-------------------------
    ' ����Ҵ� ������ ������.
    '-------------------------
    If CheckSYSTEMError(Err,True) = True Then
        Set ObjPS3G162 = Nothing		       
%>
		<Script Language=vbscript>
			parent.SetToolbar "11000001000111"
            parent.frm1.txtFromConSoNo.focus   
            parent.SetNm()
		</Script>
<%
        Response.End                                              
    End If   

    Set ObjPS3G162 = Nothing
		
	'----------------------------
	' ����Ҵ� ������ ǥ���Ѵ�.
	'----------------------------
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim iLngMaxRow       
    Dim iLngRow          
    Dim strTemp
    Dim istrData
	    
	With parent
		iLngMaxRow = .frm1.vspdData.MaxRows
		
		If "<%=ConvSPChars(EG1_exp_grp(EG1_AvaInv))%>" = "N" Then			'�������Ȯ��	
			.frm1.rdoAvaInvN.checked = True
		Else 
			.frm1.rdoAvaInvY.checked = True	
		End If
							
		If "<%=ConvSPChars(EG1_exp_grp(EG1_PurReqAuto))%>" = "N" Then	    '���ſ�û�ڵ�����	
			.frm1.rdoPurReqAutoN.checked = True
		Else
			.frm1.rdoPurReqAutoY.checked = True
		End If	
				
<%   
		For iLngRow = 0 To UBound(EG2_exp_grp,1)
    

		    If iLngRow < C_SHEETMAXROWS_D  Then
		    Else				
			   iStrNextKey = ConvSPChars(EG2_exp_grp(iLngRow, EG2_SoNo)) 			   
		       iStrNextKey = iStrNextKey & gColSep & ConvSPChars(EG2_exp_grp(iLngRow, EG2_SoSeq)) 		       
		       iStrNextKey = iStrNextKey & gColSep & ConvSPChars(EG2_exp_grp(iLngRow, EG2_SchdNo)) 		       
               Exit For
            End If 	           
%>    
			
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_ItemCd))%>"     'ǰ���ڵ�			
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_ItemName))%>"   'ǰ��� 
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_ItemSpec))%>"   '�԰� 
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_TrackingNo))%>" 'Tracking No.					
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SoUnit))%>"     '���� 
			istrData = istrData & Chr(11) &	 "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_SoQty), ggQty.DecPoint, 0)%>"			'���� 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_PreAllocQty), ggQty.DecPoint, 0)%>"		'���Ҵ緮 
			istrData = istrData & Chr(11) & ""
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_BonusQty), ggQty.DecPoint, 0)%>"		'������ 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_PreAllocBonusQty), ggQty.DecPoint, 0)%>" '���Ҵ������ 
			istrData = istrData & Chr(11) & ""
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG2_exp_grp(iLngRow, EG2_PromiseDt))%>"      '�������	
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG2_exp_grp(iLngRow, EG2_DlvyDt))%>"			'������	
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SlCd))%>"           'â���ڵ�					
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SlNm))%>"           'â��� 
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_PlantCd))%>"		'�����ڵ�			
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_PlantNm))%>"		'����� 
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SoNo))%>"			'���ֹ�ȣ			
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SoSeq))%>"			'���ּ��� 
			istrData = istrData & Chr(11) &			"<%=ConvSPChars(EG2_exp_grp(iLngRow, EG2_SchdNo))%>"		'��ǰ���� 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_PrePurReqQty), ggQty.DecPoint, 0)%>"  '�ⱸ�ſ�û��				
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG2_exp_grp(iLngRow, EG2_GiQty), ggQty.DecPoint, 0)%>"  '������				
			 
			istrData = istrData & Chr(11) & iLngMaxRow + <%=iLngRow%>			
			istrData = istrData & Chr(11) & Chr(12)		
			
<%      
		Next
%>    

		.ggoSpread.Source = .frm1.vspdData
		
		.ggoSpread.SSShowDataByClip istrData

		.lgStrPrevKey = "<%=iStrNextKey%>"
		
    	.frm1.txtHFromConSoNo.value		= "<%=ConvSPChars(Request("txtFromConSoNo"))%>"   ' Request���� hidden input���� �Ѱ��� 
		.frm1.txtHToConSoNo.value		= "<%=ConvSPChars(Request("txtToConSoNo"))%>"
		.frm1.txtHShipToParty.value		= "<%=ConvSPChars(Request("txtShipToParty"))%>"
		.frm1.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
		.frm1.txtHItem.value			= "<%=ConvSPChars(Request("txtItem"))%>"
		.frm1.txtHPlant.value			= "<%=ConvSPChars(Request("txtPlant"))%>"
		.frm1.txtHFromDate.value		= "<%=ConvSPChars(Request("txtFromDate"))%>"
		.frm1.txtHToDate.value			= "<%=ConvSPChars(Request("txtToDate"))%>"
		.frm1.txtHAllocFlagRadio.value	= "<%=ConvSPChars(Request("txtRadio"))%>"		
		
		.SetNm
		
		.DbQueryOk
	
	End With

</Script>
<%																			

Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
	
	Dim iErrorPosition
	
	Dim I1_s_inv_alloc	
	Redim I1_s_inv_alloc(1)								

	Dim txtSpread
		
    Err.Clear																		
    
	If Request("txtMaxRows") = "" Then
		Call ServerMesgBox("MaxRows ���ǰ��� ����ֽ��ϴ�!",vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
    
	I1_s_inv_alloc(0) = Trim(Request("txtHAvaInvRadio"))
	I1_s_inv_alloc(1) = Trim(Request("txtHPurReqAutoRadio"))
    txtSpread = Trim(Request("txtSpread"))
    
    Set ObjPS3G161 = Server.CreateObject("PS3G161.CsModSchdForDlvySvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If    

    Call ObjPS3G161.MODIFY_SCHDLINE_FOR_DELIVERY_SVR(gStrGlobalCollection, I1_s_inv_alloc, _
													 txtSpread, iErrorPosition)   

    If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
       Set ObjPS3G161 = Nothing
       Response.End
	End If

    Set ObjPS3G161 = Nothing 
%>
<Script Language=vbscript>
	With parent
		.SetNm																			
		.DbSaveOk
	End With
</Script>
<%

End Select
%>
