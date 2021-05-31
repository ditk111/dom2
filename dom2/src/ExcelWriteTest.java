import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



class CustomerVo {
	 
    private String  custId;        //고객ID
    private String  custName;    //고객명
    private String    custAge;    //고객나이
    private String    custEmail;    //고객이메일
    
    
    // 생성자
    public CustomerVo(String custId, String custName, String custAge,
            String custEmail) {
        super();
        this.custId = custId;
        this.custName = custName;
        this.custAge = custAge;
        this.custEmail = custEmail;
    }
    
    public String getCustId() {
        return custId;
    }
    public void setCustId(String custId) {
        this.custId = custId;
    }
    public String getCustName() {
        return custName;
    }
    public void setCustName(String custName) {
        this.custName = custName;
    }
    public String getCustAge() {
        return custAge;
    }
    public void setCustAge(String custAge) {
        this.custAge = custAge;
    }
    public String getCustEmail() {
        return custEmail;
    }
    public void setCustEmail(String custEmail) {
        this.custEmail = custEmail;
    }
    
    @Override
    public String toString() {
        StringBuffer sb = new StringBuffer();
        
        sb.append("ID : " + custId);
        sb.append(" ,NAME : " + custName);
        sb.append(" ,AGE : " + custAge);
        sb.append(" ,EMAIL : " + custEmail);
        return sb.toString();
    }
}





public class ExcelWriteTest {
    public static void main(String[] args)  {
    	
    	List<CustomerVo> list = new ArrayList<CustomerVo>();
        list.add(new CustomerVo("asdf1", "사용자1", "30", "asdf1@naver.com"));
        list.add(new CustomerVo("asdf2", "사용자2", "31", "asdf2@naver.com"));
        list.add(new CustomerVo("asdf3", "사용자3", "32", "asdf3@naver.com"));
        list.add(new CustomerVo("asdf4", "사용자4", "33", "asdf4@naver.com"));
        list.add(new CustomerVo("asdf5", "사용자5", "34", "asdf5@naver.com"));
        
        ExcelWriteTest excelWriter = new ExcelWriteTest();
        //xls 파일 쓰기
        excelWriter.xlsWiter(list);
        
        //xlsx 파일 쓰기
        excelWriter.xlsxWiter(list);

    	
    	
    	
    }
    
    public void xlsWiter(List<CustomerVo> list) {
        // 워크북 생성
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 워크시트 생성
        HSSFSheet sheet = workbook.createSheet("테스트1");
        // 행 생성
        HSSFRow row = sheet.createRow(0);
        // 쎌 생성
        HSSFCell cell;
        
        // 헤더 정보 구성
        cell = row.createCell(0);
        cell.setCellValue("아이디");
        
        cell = row.createCell(1);
        cell.setCellValue("이름");
        
        cell = row.createCell(2);
        cell.setCellValue("나이");
        
        cell = row.createCell(3);
        cell.setCellValue("이메일");
        
        // 리스트의 size 만큼 row를 생성
        CustomerVo vo;
        for(int rowIdx=0; rowIdx < list.size(); rowIdx++) {
            vo = list.get(rowIdx);
            
            // 행 생성
            row = sheet.createRow(rowIdx+1);
            
            cell = row.createCell(0);
            cell.setCellValue(vo.getCustId());
            
            cell = row.createCell(1);
            cell.setCellValue(vo.getCustName());
            
            cell = row.createCell(2);
            cell.setCellValue(vo.getCustAge());
            
            cell = row.createCell(3);
            cell.setCellValue(vo.getCustEmail());
            
        }
        
        // 입력된 내용 파일로 쓰기
        File file = new File("C:/Users/user/Desktop/개발관련자료/dom2/src/test2.xls");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(workbook!=null) workbook.close();
                if(fos!=null) fos.close();
                
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
    }
    
    public void xlsxWiter(List<CustomerVo> list) {
        // 워크북 생성
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 워크시트 생성
        XSSFSheet sheet = workbook.createSheet();
        // 행 생성
        XSSFRow row = sheet.createRow(0);
        // 쎌 생성
        XSSFCell cell;
        
        // 헤더 정보 구성
        cell = row.createCell(0);
        cell.setCellValue("아이디");
        
        cell = row.createCell(1);
        cell.setCellValue("이름");
        
        cell = row.createCell(2);
        cell.setCellValue("나이");
        
        cell = row.createCell(3);
        cell.setCellValue("이메일");
        
        // 리스트의 size 만큼 row를 생성
        CustomerVo vo;
        for(int rowIdx=0; rowIdx < list.size(); rowIdx++) {
            vo = list.get(rowIdx);
            
            // 행 생성
            row = sheet.createRow(rowIdx+1);
            
            cell = row.createCell(0);
            cell.setCellValue(vo.getCustId());
            
            cell = row.createCell(1);
            cell.setCellValue(vo.getCustName());
            
            cell = row.createCell(2);
            cell.setCellValue(vo.getCustAge());
            
            cell = row.createCell(3);
            cell.setCellValue(vo.getCustEmail());
            
        }
        
        // 입력된 내용 파일로 쓰기
        File file = new File("C:/Users/user/Desktop/개발관련자료/dom2/src/test3.xlsx");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(workbook!=null) workbook.close();
                if(fos!=null) fos.close();
                
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
    }

}

