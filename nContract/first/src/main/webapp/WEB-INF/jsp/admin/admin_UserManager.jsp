<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<%@ include file="/WEB-INF/include/include-header.jspf"%>
	<style type="text/css">
	.init_btn {	
 			border-radius: 3px;
/* 			padding: 2px 2px 2px 2px; */
			color: #fff !important;
			display: inline-block;
			background-color: #6b9ab8;
			border: 1px solid #56819d;
			vertical-align: middle;
			width: 45px;
			height: 22px;
		}
	</style>
</head>
<body>

	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>사용자 정보관리</H2>
	<br/>
<div class="grid_Layer">
	<br/>
	<table id ="grid"></table>
    <div id = "pager">
	</div>
</div>
<form id="frm" name="frm" enctype="multipart/form-data"></form>
<script type="text/javascript">
	var userId;
	var comSubmit = new ComSubmit("frm");

	$(function(){
		userId = "<sec:authentication property='principal.username'/>";
		
		$("#grid").jqGrid({
				url : "<c:url value='/admin/userInfoView.do' />",				// json객체를 반환해줄 요청URL
				caption : '사용자관리',												// caption : 그리드의 제목을 지정한다.
				datatype : 'json',												// datatype : 데이터 타입을 지정한다.
				height : '300px',												// height : 그리드의 높이를 지정한다.
				pager  : '#pager',     											// pager : 도구 모임이 될 div 태그를 지정한다.
				mtype  : 'post',                           						// mtype : 데이터 전송방식을 지정한다.
				emptyrecords : '데이터가 존재하지 않습니다.',								// 반환값이 없을 경우 화면에 출력되는 구문
				rowNum :  10, 													// 한화면에 보여줄 row의 수
				rowList : [10, 20, 30, 40],   									// 한 화면에서 볼 수 있는 row의 수를 조절 가능
				rownumbers: true,                       						// row의 숫자를 표시해준다.
				jsonReader :{
						  repeatitems : false,
		                  root : 'list',
		                  page : 'page',
		                  total : 'total',
		                  records : 'records'
				},
				
				colNames : ['ID(사번)', '이 름', '사업부', '권 한', '메일 주소', '유효구분', '비밀번호 초기화' ],
		        colModel : [
			                 { name : 'USER_NO',    index : 'USER_NO',     width : 100,  align : 'center',  editable : true, key : true},
			                 { name : 'USER_NM',    index : 'USER_NM',     width : 100,  align : 'center',  editable : true},
			                 { name : 'USER_DEPT',  index : 'USER_DEPT',   width : 80,  align : 'center',  editable : true}, //나중에 동적으로 변경해보자 (유저권한)
			                 { name : 'USER_ROLE',  index : 'USER_ROLE',   width : 80,  align : 'center',  editable : true}, //나중에 동적으로 변경해보자 (유저권한)
			                 { name : 'USER_EMAIL', index : 'USER_EMAIL',  width : 150,  align : 'center',  editable : true}, //나중에 동적으로 변경해보자 (유저권한)
	                         { name : 'VALID_FLAG', 	index : 'VALID_FLAG',  width : 80,   align : 'center',  editable : true}, //나중에 동적으로 변경해보자(사업부)
			                 { name : 'PASSWRD_INIT', index : 'PASSWRD_INIT',  width : 100,  align : 'center', formatter: fn_initBtnSet}   //나중에 동적으로 변경해보자(사용여부)
			               ],
	            gridComplete : function() {
	            },
	            loadError:	function(xhr, status, error) {
	                alert(error); 
	            }
			}).navGrid('#pager', { edit : true, add : true, del : true, search : true, },
	 	    {/* Edit options */
 	    	  	url : "<c:url value='/admin/editUserGridView.do' />",
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
			url : "<c:url value='/admin/editUserGridView.do' />",
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
			url : "<c:url value='/admin/editUserGridView.do' />",			
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
	
	});
	
	function fn_initBtnSet(cellvalue, options, rowObject) { 
		  var str = "";
		  var userKey = options.rowId.toString(); //각 user Key

		  str += "<button type='button' class='init_btn' value='" + userKey + "' onclick='fn_Passwrdinit(this)'>초기화</button>";		   
		  return str;
		}

	function fn_Passwrdinit(Obj){
		var vKey = Obj.value;
		
		Lobibox.confirm({
		    msg: "ID(사번) : [" + vKey + "] 의 비밀번호를 초기화하시겠습니까? </br> 비밀번호 id(사번)과 동일하게 초기화 됩니다.",
		    callback: function ($this, type, ev) {
		        if(type == "yes"){
// 		        	event.preventDefault();
// 		        	event.returnValue = false;
		        	
		        	fn_initSubmit(vKey);
		        }
		    }
		});
	}
	
	function fn_initSubmit(Key){
    	comSubmit.setUrl("<c:url value='/admin/initUserInfo.do'/>");
    	comSubmit.addParam("userid", userId);		//수정한 사람의 아이디
    	comSubmit.addParam("userKey", Key);			//초기화 대상이 된 아이디
    	comSubmit.addParam("type", "I");			//진행 타입 -> I: 초기화(initial), U: 수정(update) 
		
    	comSubmit.submit();
	}

</script>
	
</body>
</html>