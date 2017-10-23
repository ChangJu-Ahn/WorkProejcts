package first.common.util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
 
import javax.servlet.http.HttpServletRequest;
 
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

@Component("fileUtils")
public class FileUtils {
	//private static final String filePath = "Y:\\file\\";
	//private static final String filePath = "\\\\192.168.30.222\\102.계약서\\file\\";

	//Changed ip of new sharing folder due to break previous sharing folder.
	private static final String filePath = "\\\\192.168.10.60\\102.계약서\\file\\";
	
	public List<Map<String,Object>> parseInsertFileInfo(Map<String,Object> map, HttpServletRequest request) throws Exception{
        
		MultipartHttpServletRequest multipartHttpServletRequest = (MultipartHttpServletRequest)request;
        Iterator<String> iterator = multipartHttpServletRequest.getFileNames();
        MultipartFile multipartFile = null;
        String originalFileName = null;
        String originalFileExtension = null;
        String storedFileName = null;
         
        List<Map<String,Object>> list = new ArrayList<Map<String,Object>>();
        Map<String, Object> listMap = null; 
         
        File file = new File(filePath);
        if(file.exists() == false){
            file.mkdirs();
        }
       
        while(iterator.hasNext()){
            multipartFile = multipartHttpServletRequest.getFile(iterator.next());
            if(multipartFile.isEmpty() == false){
                originalFileName = multipartFile.getOriginalFilename();
                originalFileExtension = originalFileName.substring(originalFileName.lastIndexOf("."));
                storedFileName = CommonUtils.getRandomString() + originalFileExtension;
                
                String Contract_no = (String)map.get("CONTRACT_NO");
                String FILE_SEQ = (String)map.get("FILE_SEQ");
                String userid = (String)map.get("userid");
                
                file = new File(filePath + storedFileName);
                multipartFile.transferTo(file);
                 
                listMap = new HashMap<String,Object>();
                listMap.put("CONTRACT_NO", Contract_no);
                listMap.put("FILE_SEQ", FILE_SEQ);
                listMap.put("userid", userid);
                listMap.put("ORIGINAL_FILE_NAME", originalFileName);
                listMap.put("STORED_FILE_NAME", storedFileName);
                listMap.put("FILE_SIZE", multipartFile.getSize());
                list.add(listMap);
                
            }
        }
        return list;
    }

	public List<Map<String, Object>> parseUpdateFileInfo(Map<String, Object> map, HttpServletRequest request) throws IllegalStateException, IOException {
		
		MultipartHttpServletRequest multipartHttpServletRequest = (MultipartHttpServletRequest)request;
	    Iterator<String> iterator = multipartHttpServletRequest.getFileNames();
	     
	    MultipartFile multipartFile = null;
	    String originalFileName = null;
	    String originalFileExtension = null;
	    String storedFileName = null;
	     
	    List<Map<String,Object>> list = new ArrayList<Map<String,Object>>();
	    Map<String, Object> listMap = null; 
	     
	    String Contract_no = (String)map.get("CONTRACT_NO");
	    String FILE_SEQ = (String)map.get("FILE_SEQ");
	    String userid = (String)map.get("userid");
	    String requestName = null;
	    String idx = null;
	    
	     
	    while(iterator.hasNext()){
	        multipartFile = multipartHttpServletRequest.getFile(iterator.next());
	        if(multipartFile.isEmpty() == false){
	            originalFileName = multipartFile.getOriginalFilename();
	            originalFileExtension = originalFileName.substring(originalFileName.lastIndexOf("."));
	            storedFileName = CommonUtils.getRandomString() + originalFileExtension;
	             
	            multipartFile.transferTo(new File(filePath + storedFileName));
	             
	            listMap = new HashMap<String,Object>();
	            listMap.put("IS_NEW", "Y");
	            listMap.put("FILE_SEQ", FILE_SEQ);
	            listMap.put("CONTRACT_NO", Contract_no);
	            listMap.put("ORIGINAL_FILE_NAME", originalFileName);
	            listMap.put("STORED_FILE_NAME", storedFileName);
	            listMap.put("userid", userid);
	            listMap.put("FILE_SIZE", multipartFile.getSize());
	            list.add(listMap);
	            	            
	        }
	        else{
	            requestName = multipartFile.getName();
	            idx = "IDX_"+requestName.substring(requestName.indexOf("_")+1);
	            if(map.containsKey(idx) == true && map.get(idx) != null){
	                listMap = new HashMap<String,Object>();
	                listMap.put("IS_NEW", "N");
	                listMap.put("CONTRACT_NO", Contract_no);
	                listMap.put("userid", userid);
	                listMap.put("SEQ", map.get(idx));
	                listMap.put("FILE_SEQ", FILE_SEQ);
	                list.add(listMap);
	            }
	        }
	    }
	    return list;
	}

}
