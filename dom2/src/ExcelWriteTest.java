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
	 
    private String  custId;        //��ID
    private String  custName;    //����
    private String    custAge;    //������
    private String    custEmail;    //���̸���
    
    
    // ������
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
        list.add(new CustomerVo("asdf1", "�����1", "30", "asdf1@naver.com"));
        list.add(new CustomerVo("asdf2", "�����2", "31", "asdf2@naver.com"));
        list.add(new CustomerVo("asdf3", "�����3", "32", "asdf3@naver.com"));
        list.add(new CustomerVo("asdf4", "�����4", "33", "asdf4@naver.com"));
        list.add(new CustomerVo("asdf5", "�����5", "34", "asdf5@naver.com"));
        
        ExcelWriteTest excelWriter = new ExcelWriteTest();
        //xls ���� ����
        excelWriter.xlsWiter(list);
        
        //xlsx ���� ����
        excelWriter.xlsxWiter(list);

    	
    	
    	
    }
    
    public void xlsWiter(List<CustomerVo> list) {
        // ��ũ�� ����
        HSSFWorkbook workbook = new HSSFWorkbook();
        // ��ũ��Ʈ ����
        HSSFSheet sheet = workbook.createSheet("�׽�Ʈ1");
        // �� ����
        HSSFRow row = sheet.createRow(0);
        // �� ����
        HSSFCell cell;
        
        // ��� ���� ����
        cell = row.createCell(0);
        cell.setCellValue("���̵�");
        
        cell = row.createCell(1);
        cell.setCellValue("�̸�");
        
        cell = row.createCell(2);
        cell.setCellValue("����");
        
        cell = row.createCell(3);
        cell.setCellValue("�̸���");
        
        // ����Ʈ�� size ��ŭ row�� ����
        CustomerVo vo;
        for(int rowIdx=0; rowIdx < list.size(); rowIdx++) {
            vo = list.get(rowIdx);
            
            // �� ����
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
        
        // �Էµ� ���� ���Ϸ� ����
        File file = new File("C:/Users/user/Desktop/���߰����ڷ�/dom2/src/test2.xls");
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
        // ��ũ�� ����
        XSSFWorkbook workbook = new XSSFWorkbook();
        // ��ũ��Ʈ ����
        XSSFSheet sheet = workbook.createSheet();
        // �� ����
        XSSFRow row = sheet.createRow(0);
        // �� ����
        XSSFCell cell;
        
        // ��� ���� ����
        cell = row.createCell(0);
        cell.setCellValue("���̵�");
        
        cell = row.createCell(1);
        cell.setCellValue("�̸�");
        
        cell = row.createCell(2);
        cell.setCellValue("����");
        
        cell = row.createCell(3);
        cell.setCellValue("�̸���");
        
        // ����Ʈ�� size ��ŭ row�� ����
        CustomerVo vo;
        for(int rowIdx=0; rowIdx < list.size(); rowIdx++) {
            vo = list.get(rowIdx);
            
            // �� ����
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
        
        // �Էµ� ���� ���Ϸ� ����
        File file = new File("C:/Users/user/Desktop/���߰����ڷ�/dom2/src/test3.xlsx");
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

