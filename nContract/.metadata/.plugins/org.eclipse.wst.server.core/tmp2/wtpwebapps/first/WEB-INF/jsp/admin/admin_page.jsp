<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
<style type="text/css">
    /* 이거 왜 CSS에 넣으면 안되지 ㅠㅠ */
 	.layer
	{
	  position:absolute;
	  top:50%;
	  left:50%;  
	  transform:translate(-50%, -50%)
	} 
</style>

</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	      
<h2>기준정보 입력 및 수정</h2>
<br/>

<div class="layer">
	<select id = "gubun">
		<option value="A">계약구분</option>
		<option value="B">사업부</option>
		<option value="C">구분</option>
		<option value="D">해지조건</option>
		<option value="E">거래처</option>
		<option value="F">자동연장</option>
		<option value="G">해지통지기간</option>
		<option value="H">메일참조자</option>
	</select>
	<input type="button" id ="btn" value = "검색"/>
	<br/>
	<table id ="grid"></table>
    <div id = "pager">
	</div>
    <div id = "excel" style="float:right;margin-top: 10px;"> <!-- 2016.10.25, ahncj : 관리자가 기준정보를 엑셀로 내려받아 사용할 수 있도록 추가-->
    	 <img id="excelBtnImg" src='../images/btn_excelDown.gif'/>
    </div>
</div>	

<%@ include file="/WEB-INF/include/include-body.jspf" %>
<script type="text/javascript">
     $(function(){   
    	 var rowId = "";
    	 var userId = "<sec:authentication property='principal.username'/>";    	 
    	 $("#grid").jqGrid({
 		  	 url : "<c:url value='/admin/gridView.do' />",
 	         caption : '기준정보',    		    		// caption : 그리드의 제목을 지정한다.
 	         datatype : 'json',              	        // datatype : 데이터 타입을 지정한다.
 	         //ajaxGridOptions: { contentType: "application/json;charset=EUC-KR" },
 	         postData : { CODE : 'A' },                                           
 	         mtype  : 'post',                           // mtype : 데이터 전송방식을 지정한다.
 	         height : '300px',                          // height : 그리드의 높이를 지정한다.
 	         emptyrecords : '데이터가 존재하지 않습니다.',
 	         pager  : '#pager',                         // pager : 도구 모임이 될 div 태그를 지정한다.
 	         rowNum :  10,                              // rowNum : 한 화면에 표시할 행 개수를 지정한다.
 	         //loadonce : true,                         // loadonce : rowNum 설정을 사용하기 위해서 true로 지정한다.
 	         rowList : [10, 20, 30, 40],                // rowList : rowNum을 선택할 수 있는 옵션을 지정한다.	         
 	         jsonReader : {
                  repeatitems : false,
                  id : 'CODE',
                  root : 'list',
                  page : 'page',
                  total : 'total',
                  records : 'records'
  			 },
 	         // colNames : 열의 이름을 지정한다.
 	         colNames : [ 'CODE', 'CODE 명', '약자', '상위 CODE', 'LEVEL' ],
 	         colModel : [
 	                     { name : 'CODE',          index : 'CODE',          width : 100,     align : 'center',   key : true},
 	                     { name : 'CODE_NM',       index : 'CODE_NM',       width : 300,     align : 'center',   editable : true},
 	                     { name : 'CODE_SNM',      index : 'CODE_SNM',      width : 100,     align : 'center',   editable : true},
 	                     { name : 'HIGH_CODE',     index : 'HIGH_CODE',     width : 100,     align : 'center',   editable : true},
 	                     { name : 'LVL',     	   index : 'LVL',           width : 70,      align : 'center'}
 	                    ],
 	         gridComplete : function() {
             },
             loadError:function(xhr, status, error) {
              // 데이터 로드 실패시 실행되는 부분
                 alert(error); 
             }, 
             /* onSelectRow: function (ids) {
            	//rowId  = $(“#jGrid”).jqGrid(‘getGridParam’, “selrow” );			
            	//tmp_actId = $("#grid").jqGrid("getRowData", ids).CODE_NM;		    // 선택한 열중에서 CODE_NM를 가져온다. 
     			
            	rowId = ids;
     			console.log(rowId);
     		} */
 	       // navGrid() 메서드는 검색 및 기타기능을 사용하기위해 사용된다.
 	       }).navGrid('#pager', { edit : true, add : true, del : true, search : true, },
	 	    {/* Edit options */
 	    	  	url : "<c:url value='/admin/editGridView.do' />",
 	    	  	closeAfterEdit : true,
 	    	  	reloadAfterSubmit : true,				
				serializeEditData: function (data) {
					console.log(data);
                    var postData = {s_data : data,
                    				userid : userId     
                    			   };
                    return JSON.stringify(postData);
               },
			   afterSubmit : function(response) {
				   if (response.responseText != "") {
					   Lobibox.alert("error", 
					    		 {
					    		     msg: "수정 실패"
					    		 });
					   return [ false, "Error" ];
					} else {
						Lobibox.alert("success", 
					    		 {
					    		     msg: "수정 완료"
					    		 });
					   return [ true,  "Ok" ];
			   }
			}
		}, {/* Add options */
			url : "<c:url value='/admin/editGridView.do' />",
			closeAfterAdd : true,
			reloadAfterSubmit : true,			
			serializeEditData: function (data) {
				console.log(data);
                var postData = {s_data : data,
                				userid : userId     
                			   };
                return JSON.stringify(postData);
            },
		   afterSubmit : function(response) {
			   if (response.responseText != "") {
				   Lobibox.alert("error", 
				    		 {
				    		     msg: "추가 실패"
				    		 });
				   return [ false, "Error" ];
				} else {
					Lobibox.alert("success", 
				    		 {
				    		     msg: "추가 완료"
				    		 });
				   return [ true,  "Ok" ];
		   	}
		  }
		}, {/* Delete options */
			url : "<c:url value='/admin/editGridView.do' />",			
			reloadAfterSubmit : true,
			serializeDelData : function(data) {
				console.log(data);
                var postData = {s_data : data,
                				userid : userId     
                			   };
                return JSON.stringify(postData);
            },
		   afterSubmit : function(response) {
			   if (response.responseText != "") {
				   Lobibox.alert("error", 
				    		 {
				    		     msg: "삭제 실패"
				    		 });
				   return [ false, "Error" ];
				} else {
					Lobibox.alert("success", 
				    		 {
				    		     msg: "삭제 완료"
				    		 });
				   return [ true,  "Ok" ];
		   	}
		  }
		}, {/* Search options */});

		$("body").on("click", '#btn', function(e) { //글쓰기 버튼
			e.preventDefault();
			var selectedData = $('#gubun option:selected').val();
			
			fn_openadminCode(selectedData);
		});
	});
     
 	$(document).ready(function(){
		//If this buttton push, it can download the site standard information. 
		$("#excel").click(function(e){
			e.preventDefault();
			fn_downAdminListExcelFile();
		});
    });
 
    function fn_downAdminListExcelFile(){
    	var selectedData = $('#gubun option:selected').val();	//Set the Selected Dropdown list value     	
    	var comSubmit = new ComSubmit();
    	var tempTarget = "ADMINCODE_LIST";
        
     	comSubmit.setUrl("<c:url value='/downExcel.do' />");
        comSubmit.addParam("CODE", selectedData);
        comSubmit.addParam("TARGET", tempTarget);
        
        comSubmit.submit();
    }
     
 	function fn_openadminCode(selectedData) {
		$("#grid").setGridParam({
			postData : {
				CODE : selectedData
			},
			datatype : 'json',
			page : 1
		}).trigger('reloadGrid');
	}
</script>
</body>
</html>