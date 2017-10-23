

'===================================================================
'페이지마다 통일토록 함.
'gbn:popup 구분 
'e : 포커스할 객체 
'lws 
'===================================================================

Function OpenPopupw(ByVal gbn,ByVal e)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
    Dim iCalledAspName
    Dim IntRetCD
    on Error Resume Next
 
 //품목구분시에는 품목계정이 먼저 선행되어야함 
	if gbn="item_kind" then 
		if frm1.txtItem_acct.value ="" then
		 Call DisplayMsgBox("800489","X","품목계정","X")
		 frm1.txtItem_acct.focus()
		 Exit Function
		end if
	
	end if
	
	If IsOpenPop = True Then Exit Function		 
	IsOpenPop = True
	
	
	 '-------------------------------------------
     if gbn="item_cd" then '====품목코드 
      '-------------------------------------------
		arrParam(0) = "품목코드 POPUP"						'' 팝업 명칭 %>
		arrParam(1) = "B_CIS_ITEM_MASTER"						'<%' TABLE 명칭 %>
		arrParam(2) = eval("frm1."+e).value 	'<%' Code Condition%>
		arrParam(4) = " "	'<%' Where Condition%>
		arrParam(5) = "품목코드"					'<%' 조건필드의 라벨 명칭 %>
		arrParam(3) = ""							'	<%' Name Cindition%>
		arrField(0) = "ITEM_CD"						'<%' Field명(0)%>
		arrField(1) = "ITEM_NM"					'<%' Field명(1)%>
		arrHeader(0) = "품목코드"							'<%' Header명(0)%>
		arrHeader(1) = "품목명"							'<%' Header명(1)%>
     '-------------------------------------------
     elseif  gbn="item_acct" then '====품목계정 
      '-------------------------------------------
		arrParam(0) = "품목계정 POPUP"						
		arrParam(1) = "B_MINOR"						
		arrParam(2) = eval("frm1."+e).value 
		arrParam(4) = " MAJOR_CD = N'P1001' "	
		arrParam(5) = "품목계정"					
		arrParam(3) = ""								
		arrField(0) = "MINOR_CD"						
		arrField(1) = "MINOR_NM"					
		arrHeader(0) = "품목계정"							
		arrHeader(1) = "품목계정명"							
	 '-------------------------------------------	
     elseif  gbn="item_kind" then '====품목구분 
     '-------------------------------------------
  		
		arrParam(0) = "품목구분 POPUP"						
		arrParam(1) = "B_MINOR A, B_CIS_CONFIG B "						
		arrParam(2) = eval("frm1."+e).value 
		arrParam(4) = " MAJOR_CD = N'Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = "&filtervar(frm1.txtitem_acct.value,"''","S")&" "	
		arrParam(5) = "품목구분"					
		arrParam(3) = ""								
		arrField(0) = "MINOR_CD"						
		arrField(1) = "MINOR_NM"					
		arrHeader(0) = "품목구분"							
		arrHeader(1) = "품목구분명"	
		
								
	 '-------------------------------------------	
     elseif  gbn="user" then '====User
     '-------------------------------------------
		arrParam(0) = "사용자정보 POPUP"						
		arrParam(1) = "Z_USR_MAST_REC"	
		arrParam(2) = eval("frm1."+e).value 	
		arrParam(3) = ""								
		arrParam(4) = " "	
		arrParam(5) = "사용자 ID"					
		arrField(0) = "USR_ID"						
		arrField(1) = "USR_NM"					
		arrHeader(0) = "사용자 ID"							
		arrHeader(1) = "사용자명"		
								
	 '-------------------------------------------	
     elseif  gbn="req_user" then '====req_user
     '-------------------------------------------
		arrParam(0) = "의뢰자 POPUP"						
		arrParam(1) = "B_MINOR"	
		arrParam(2) = eval("frm1."+e).value 	
		arrParam(3) = ""								
		arrParam(4) = " MAJOR_CD = N'Y1006' "
		arrParam(5) = "의뢰자"					
		arrField(0) = "MINOR_CD"						
		arrField(1) = "MINOR_NM"					
		arrHeader(0) = "의뢰자"							
		arrHeader(1) = "의뢰자명"		
			
							
     end if
     
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItem_kind.focus
		Exit Function
	Else

		eval("frm1."+e).focus
		eval("frm1."+e).value			= arrRet(0)  
		eval("frm1."+e+"_nm").Value     = arrRet(1)  
		//eval("frm1."+e).focus
		Set gActiveElement = document.activeElement
	End If	
End Function



'========================================================================================
' Function Name : txtitem_acct_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtitem_acct_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtitem_acct.value = "" Then
        frm1.txtitem_acct_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='P1001' and minor_cd="&filterVar(frm1.txtitem_acct.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtitem_acct_nm.value=""
        Else
            frm1.txtitem_acct_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    call txtItem_kind_OnChange()
End Function


'========================================================================================
' Function Name : txtItem_kind_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_kind_OnChange()
    Dim iDx
    Dim IntRetCd
 
	
    If frm1.txtItem_kind.value = "" Then
        frm1.txtItem_kind_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm ","  B_MINOR A, B_CIS_CONFIG B "," major_cd='Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = "&filtervar(frm1.txtitem_acct.value,"''","S")&" and minor_cd="&filterVar(frm1.txtItem_kind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_kind_nm.value=""
        Else
            frm1.txtItem_kind_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function





'========================================================================================
' Function Name : txtItem_lvl1_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_lvl1_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_lvl1.value = "" Then
        frm1.txtItem_lvl1_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" CLASS_NAME "," B_CIS_ITEM_CLASS "," ITEM_ACCT="&filterVar(frm1.txtitem_acct.value,"''","S") & " AND ITEM_KIND="&filterVar(frm1.txtitem_kind.value,"''","S") & " AND CLASS_CD="&filterVar(frm1.txtItem_lvl1.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_lvl1_nm.value=""
        Else
            frm1.txtItem_lvl1_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


'========================================================================================
' Function Name : txtItem_lvl2_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_lvl2_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_lvl2.value = "" Then
        frm1.txtItem_lvl2_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" CLASS_NAME "," B_CIS_ITEM_CLASS "," ITEM_ACCT="&filterVar(frm1.txtitem_acct.value,"''","S") & " AND ITEM_KIND="&filterVar(frm1.txtitem_kind.value,"''","S") & " AND PARENT_CLASS_CD="&filterVar(frm1.txtItem_lvl1.value,"''","S") & " AND CLASS_CD="&filterVar(frm1.txtItem_lvl2.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_lvl2_nm.value=""
        Else
            frm1.txtItem_lvl2_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


'========================================================================================
' Function Name : txtItem_lvl3_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_lvl3_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_lvl3.value = "" Then
        frm1.txtItem_lvl3_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" CLASS_NAME "," B_CIS_ITEM_CLASS "," ITEM_ACCT="&filterVar(frm1.txtitem_acct.value,"''","S") & " AND ITEM_KIND="&filterVar(frm1.txtitem_kind.value,"''","S") & " AND PARENT_CLASS_CD="&filterVar(frm1.txtItem_lvl2.value,"''","S") & " AND CLASS_CD="&filterVar(frm1.txtItem_lvl3.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_lvl3_nm.value=""
        Else
            frm1.txtItem_lvl3_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function



'========================================================================================
' Function Name : txtItem_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_cd_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtItem_cd.value = "" Then
        frm1.txtItem_cd_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" ITEM_NM "," B_CIS_ITEM_MASTER ","  ITEM_CD="&filterVar(frm1.txtItem_cd.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_cd_nm.value=""
        Else
            frm1.txtItem_cd_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'========================================================================================
' Function Name : txtreq_user_OnChange
' Function Desc : 
'========================================================================================
Function txtreq_user_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtreq_user.value = "" Then
        frm1.txtreq_user_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1006' and minor_cd="&filterVar(frm1.txtreq_user.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtreq_user_nm.value=""
        Else
            frm1.txtreq_user_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'========================================================================================
' Function Name : txtPurVendor_OnChange
' Function Desc : 
'========================================================================================
Function txtPurVendor_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtPurVendor.value = "" Then
        frm1.txtPurVendornm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","  BP_TYPE In ('S','CS') And usage_flag='Y' AND IN_OUT_FLAG = 'O' and BP_CD="&filterVar(frm1.txtPurVendor.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtPurVendornm.value=""
        Else
            frm1.txtPurVendornm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'========================================================================================
' Function Name : txtHSCd_OnChange
' Function Desc : 
'========================================================================================
Function txtHSCd_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtHSCd.value = "" Then
        frm1.txtHSnm.value = ""
    ELSE    
		IntRetCd =  CommonQueryRs(" HS_NM "," B_HS_CODE "," HS_CD="&filterVar(frm1.txtHSCd.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)      
		  If IntRetCd = false Then
			 frm1.txtHSnm.value=""
        Else
            frm1.txtHSnm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'========================================================================================
' Function Name : txtPurGroup_OnChange
' Function Desc : 
'========================================================================================
Function txtPurGroup_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtPurGroup.value = "" Then
        frm1.txtPurGroupnm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" PUR_GRP_NM "," B_Pur_Grp ","  USAGE_FLG='Y' AND PUR_GRP="&filterVar(frm1.txtPurGroup.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtPurGroupnm.value=""
        Else
            frm1.txtPurGroupnm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

