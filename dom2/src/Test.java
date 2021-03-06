import java.awt.FlowLayout;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


class Student{
	ArrayList<String> origin_date=new ArrayList<>();
	ArrayList<Integer> date=new ArrayList<>();
	ArrayList<Integer> origin_time=new ArrayList<>();
	ArrayList<Integer> time=new ArrayList<>();
	ArrayList<String> origin_machine=new ArrayList<>();
	ArrayList<String> machine=new ArrayList<>();
	ArrayList<String> origin_console=new ArrayList<>();
	ArrayList<String> console=new ArrayList<>();
	ArrayList<String> card=new ArrayList<>();
	ArrayList<String> number=new ArrayList<>();
	ArrayList<String> username=new ArrayList<>();
	ArrayList<String> company=new ArrayList<>();
	ArrayList<String> dept=new ArrayList<>();
	ArrayList<String> etc=new ArrayList<>();
	ArrayList<String> because=new ArrayList<>();
	ArrayList<Integer> tag=new ArrayList<>();
	int penalty=0;
	boolean caution=false;
}

class xlsx{			
public static String[][] read(String filepath) {
		try {
			FileInputStream file = new FileInputStream(filepath);
	        XSSFWorkbook xworkbook = new XSSFWorkbook(file);
	           
	        XSSFSheet sheet=xworkbook.getSheetAt(0);   //???? ?? (?????????? ?????????? 0?? ????) / ???? ?? ?????? ???????????? FOR???? ?????? ????????
	        String value="";   // ???? ?? ??????
	        int rowindex=0;    // ?? ?????? ??????
	        int columnindex=0; // ?? ?????? ??????
	        int rows=sheet.getPhysicalNumberOfRows();// ?? ???? ??
	            
	        XSSFRow firstrow=sheet.getRow(rowindex); //???? ????????
	        int firstcells=firstrow.getPhysicalNumberOfCells(); // ?????? ???? ?? ??????
	            
	        String[][] human=new String[rows][firstcells]; //2???????? ??????

	        for(rowindex=0;rowindex<rows;rowindex++){
	            XSSFRow row=sheet.getRow(rowindex); //?? ????????
	            if(row !=null){
	                int cells=row.getPhysicalNumberOfCells(); // ???? ?? ????????
	                for(columnindex=0; columnindex<=cells; columnindex++){
	                    //?????? ??????
	                    XSSFCell cell=row.getCell(columnindex);
	                    /*String value="";
	                    //???? ???????????? ???? ??????*/
	                    if(cell==null){
	                        continue;
	                    }else{
	                        //???????? ???? ????
	                        switch (cell.getCellType()){
	                        case FORMULA:
	                            value=cell.getCellFormula();
	                            break;
	                        case NUMERIC:
	                            value=cell.getNumericCellValue()+"";
	                            break;
	                        case STRING: // ???????? String ???? ????????. ?????? String???? ????????.
	                            value=cell.getStringCellValue()+"";
	                            break;
	                        case BLANK:
	                            value=cell.getBooleanCellValue()+"";
	                            break;
	                        case ERROR:
	                            value=cell.getErrorCellValue()+"";
	                            break;
	                        }
	                    }
	                    human[rowindex][columnindex]=value;
	                }
	            }
	        }
	        return human;
 
	}catch(Exception e) {
		e.printStackTrace();
	}
	return null;
	}

public static void write(Student[] student,String file_save_path,int date) {
	
		 // ?????? ????
        XSSFWorkbook xworkbook = new XSSFWorkbook();
        
        // ???????? ????
        XSSFSheet sheet = xworkbook.createSheet();
        // ?? ????
        XSSFRow row=sheet.createRow(1);
        // ?? ????
        XSSFCell cell=row.createCell(0);
        
        //title
        XSSFFont TitleFont=xworkbook.createFont();
        TitleFont.setFontHeightInPoints((short)13);
        TitleFont.setFontName("???? ????");
        TitleFont.setColor(IndexedColors.BLUE.getIndex());
        TitleFont.setBold(true);
        
        CellStyle TitleStyle=xworkbook.createCellStyle();
        TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        TitleStyle.setFont(TitleFont);
        
        //Head
        XSSFFont HeadFont=xworkbook.createFont();
        HeadFont.setBold(true);
        HeadFont.setFontHeightInPoints((short)13);
        HeadFont.setFontName("???? ????");
        
        CellStyle HeadStyle=xworkbook.createCellStyle();
        HeadStyle.setAlignment(HorizontalAlignment.CENTER);
        HeadStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        HeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HeadStyle.setBorderTop(BorderStyle.THIN);
        HeadStyle.setBorderBottom(BorderStyle.THIN);
        HeadStyle.setBorderLeft(BorderStyle.THIN);
        HeadStyle.setBorderRight(BorderStyle.THIN);
        HeadStyle.setFont(HeadFont);
        
        //Body
        XSSFFont BodyFont=xworkbook.createFont();
        BodyFont.setFontHeightInPoints((short)11);
        BodyFont.setFontName("???? ????");
        
        CellStyle BodyStyle=xworkbook.createCellStyle();
        BodyStyle.setAlignment(HorizontalAlignment.CENTER);
        BodyStyle.setBorderTop(BorderStyle.THIN);
        BodyStyle.setBorderBottom(BorderStyle.THIN);
        BodyStyle.setBorderLeft(BorderStyle.THIN);
        BodyStyle.setBorderRight(BorderStyle.THIN);
        BodyStyle.setFont(BodyFont);
        
        // ????
        sheet.setColumnWidth(0, (short)8000); //????
        sheet.setColumnWidth(1, (short)3500); //??????
        sheet.setColumnWidth(2, (short)6500); //??????
        sheet.setColumnWidth(3, (short)3000); //???? ??????
        sheet.setColumnWidth(4, (short)3000); // ????
        sheet.setColumnWidth(5, (short)3500); // ?????? ??
        sheet.setColumnWidth(6, (short)3000); // ????????
        sheet.setColumnWidth(7, (short)3000); // ????????
        sheet.setColumnWidth(8, (short)3000); // ????

        int count=1; // count?? ?? ?????? ??????
        

        count=print_simple(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ??????

        count++;
        count=print_detail(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ?????? ????????
        
        count++;
        count=print_simple(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ??????
                
        count++;
        count=print_detail(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ?????? ????????
        
        // ?????? ???? ?????? ???? - ???? ???? ?????? ????
        File file = new File(file_save_path+date+" ????????.xlsx");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            xworkbook.write(fos);
            
			System.out.println();
			System.out.println("--------------- ???? ???? ???????? ????. ---------------");
			System.out.println();
			
        } catch (FileNotFoundException e) {
            System.out.println("???? : student ???? ?????? ???? ???? ????????????.");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(xworkbook!=null) xworkbook.close();
                if(fos!=null) fos.close();
                
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
		
		
	}

public static int print_simple(int count,int penalty,XSSFSheet sheet,XSSFRow row,XSSFCell cell,CellStyle TitleStyle,CellStyle HeadStyle,CellStyle BodyStyle,Student[] student) {
	
	row = sheet.createRow(count);
    cell=row.createCell(0);
    if(penalty>=3) {
    	cell.setCellValue("???? 3?????? ??????");
    }
    else {
    	cell.setCellValue("???? 3?????? ??????");	
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    cell=row.createCell(0);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(1);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(2);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    count++;
    
    int a = 0,b=0;
    
    if(penalty==3) {
    	a=3;
    	b=20;    	
    }
    else if(penalty==1) {
    	a=1;
    	b=3;
    }
    
    for(int i=0;i<student.length;i++) {
    	for(int j=0;j<student[i].date.size();j++) {
    		if(student[i].penalty>=a&&student[i].penalty<b) { // if???? ???? ???????? ?????????? ?? ???????????????
    			row=sheet.createRow(count);
    			
    			cell = row.createCell(0);
    			cell.setCellValue(student[i].number.get(j));
    	        cell.setCellStyle(BodyStyle);
    			
                cell = row.createCell(1);
    			cell.setCellValue(student[i].username.get(j));
    	        cell.setCellStyle(BodyStyle);
    	        
    	        cell=row.createCell(2);
    			String s="";
    			for(int k=0;k<student[i].because.size();k++) {
    				s+=student[i].because.get(k)+" ";
    			}
    			cell.setCellValue(s);
    	        cell.setCellStyle(BodyStyle);
    			
    			cell = row.createCell(3);
    			cell.setCellValue(student[i].penalty+" ??");
    	        cell.setCellStyle(BodyStyle);
    			
    			count++;
    			break;
    		}
    		
    	}
    }
	
	return count;
}

public static int print_detail(int count,int penalty,XSSFSheet sheet,XSSFRow row,XSSFCell cell,CellStyle TitleStyle,CellStyle HeadStyle,CellStyle BodyStyle,Student[] student) {
	
	row=sheet.createRow(count);
    
    cell=row.createCell(0);
    if(penalty>=3) {
    	cell.setCellValue("???? 3?????? ?????? ????????");
    }
    else {
    	cell.setCellValue("???? 3?????? ?????? ????????");
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    
    // ???? ???? ????
    cell=row.createCell(0);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    
    cell=row.createCell(1);
    cell.setCellValue("???? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(2);
    cell.setCellValue("???? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("???? ??????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(4);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(5);
    cell.setCellValue("?????? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(6);
    cell.setCellValue("???? ????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(7);
    cell.setCellValue("???? ????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(8);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    count++;
    
    int a=0,b=0;
    
    if(penalty==3) {
    	a=3;
    	b=20;
    }
    else if(penalty==1) {
    	a=1;
    	b=3;
    }
    

    Student vo;
    for(int rowIdx=0; rowIdx < student.length; rowIdx++) {
        vo = student[rowIdx];
    	for(int j=0;j<vo.tag.size();j++) {
    		int k=vo.tag.get(j);
    		
            if(student[rowIdx].penalty>=a&&student[rowIdx].penalty<b) {
            	
        	// ?? ????
            row = sheet.createRow(count++);
            
            cell = row.createCell(0);
            cell.setCellValue(vo.origin_date.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(1);
            cell.setCellValue(vo.origin_machine.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(2);
            cell.setCellValue(vo.origin_console.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(3);
            cell.setCellValue(vo.card.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(4);
            cell.setCellValue(vo.number.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(5);
            cell.setCellValue(vo.username.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(6);
            cell.setCellValue(vo.company.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(7);
            cell.setCellValue(vo.dept.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(8);
            cell.setCellValue(vo.etc.get(k));
	        cell.setCellStyle(BodyStyle);
            
            }
    	}
    }
	
	return count;
}

}


class xls{			
public static String[][] read(String filepath) {
		try {
			FileInputStream file = new FileInputStream(filepath);
	        HSSFWorkbook hworkbook = new HSSFWorkbook(file);
	           
	        HSSFSheet sheet=hworkbook.getSheetAt(0);   //???? ?? (?????????? ?????????? 0?? ????) / ???? ?? ?????? ???????????? FOR???? ?????? ????????
	        String value="";   // ???? ?? ??????
	        int rowindex=0;    // ?? ?????? ??????
	        int columnindex=0; // ?? ?????? ??????
	        int rows=sheet.getPhysicalNumberOfRows();// ?? ???? ??
	            
	        HSSFRow firstrow=sheet.getRow(rowindex); //???? ????????
	        int firstcells=firstrow.getPhysicalNumberOfCells(); // ?????? ???? ?? ??????
	            
	        String[][] human=new String[rows][firstcells]; //2???????? ??????

	        for(rowindex=0;rowindex<rows;rowindex++){
	            HSSFRow row=sheet.getRow(rowindex); //?? ????????
	            if(row !=null){
	                int cells=row.getPhysicalNumberOfCells(); // ???? ?? ????????
	                for(columnindex=0; columnindex<=cells; columnindex++){
	                    //?????? ??????
	                    HSSFCell cell=row.getCell(columnindex);
	                    /*String value="";
	                    //???? ???????????? ???? ??????*/
	                    if(cell==null){
	                        continue;
	                    }else{
	                        //???????? ???? ????
	                        switch (cell.getCellType()){
	                        case FORMULA:
	                            value=cell.getCellFormula();
	                            break;
	                        case NUMERIC:
	                            value=cell.getNumericCellValue()+"";
	                            break;
	                        case STRING: // ???????? String ???? ????????. ?????? String???? ????????.
	                            value=cell.getStringCellValue()+"";
	                            break;
	                        case BLANK:
	                            value=cell.getBooleanCellValue()+"";
	                            break;
	                        case ERROR:
	                            value=cell.getErrorCellValue()+"";
	                            break;
	                        }
	                    }
	                    human[rowindex][columnindex]=value;
	                }
	            }
	        }
	        return human;
 
	}catch(Exception e) {
		e.printStackTrace();
	}
	return null;
	}

public static void write(Student[] student,String file_save_path,int date) {
	
		 // ?????? ????
        HSSFWorkbook hworkbook = new HSSFWorkbook();
        
        // ???????? ????
        HSSFSheet sheet = hworkbook.createSheet();
        // ?? ????
        HSSFRow row=sheet.createRow(1);
        // ?? ????
        HSSFCell cell=row.createCell(0);
        
        //title
        HSSFFont TitleFont=hworkbook.createFont();
        TitleFont.setFontHeightInPoints((short)13);
        TitleFont.setFontName("???? ????");
        TitleFont.setColor(IndexedColors.BLUE.getIndex());
        TitleFont.setBold(true);
        
        CellStyle TitleStyle=hworkbook.createCellStyle();
        TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        TitleStyle.setFont(TitleFont);
        
        //Head
        HSSFFont HeadFont=hworkbook.createFont();
        HeadFont.setBold(true);
        HeadFont.setFontHeightInPoints((short)13);
        HeadFont.setFontName("???? ????");
        
        CellStyle HeadStyle=hworkbook.createCellStyle();
        HeadStyle.setAlignment(HorizontalAlignment.CENTER);
        HeadStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        HeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HeadStyle.setBorderTop(BorderStyle.THIN);
        HeadStyle.setBorderBottom(BorderStyle.THIN);
        HeadStyle.setBorderLeft(BorderStyle.THIN);
        HeadStyle.setBorderRight(BorderStyle.THIN);
        HeadStyle.setFont(HeadFont);
        
        //Body
        HSSFFont BodyFont=hworkbook.createFont();
        BodyFont.setFontHeightInPoints((short)11);
        BodyFont.setFontName("???? ????");
        
        CellStyle BodyStyle=hworkbook.createCellStyle();
        BodyStyle.setAlignment(HorizontalAlignment.CENTER);
        BodyStyle.setBorderTop(BorderStyle.THIN);
        BodyStyle.setBorderBottom(BorderStyle.THIN);
        BodyStyle.setBorderLeft(BorderStyle.THIN);
        BodyStyle.setBorderRight(BorderStyle.THIN);
        BodyStyle.setFont(BodyFont);
        
        // ????
        sheet.setColumnWidth(0, (short)10000); //????
        sheet.setColumnWidth(1, (short)4000); //??????
        sheet.setColumnWidth(2, (short)8000); //??????
        sheet.setColumnWidth(3, (short)4000); //???? ??????
        sheet.setColumnWidth(4, (short)3000); // ????
        sheet.setColumnWidth(5, (short)4000); // ?????? ??
        sheet.setColumnWidth(6, (short)3000); // ????????
        sheet.setColumnWidth(7, (short)3000); // ????????
        sheet.setColumnWidth(8, (short)3000); // ????

        int count=1; // count?? ?? ?????? ??????
        

        count=print_simple(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ??????

        count++;
        count=print_detail(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ?????? ????????
        
        count++;
        count=print_simple(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ??????
                
        count++;
        count=print_detail(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// ???? 3?????? ?????? ????????
        
        // ?????? ???? ?????? ???? - ???? ???? ?????? ????
        File file = new File(file_save_path+date+" ????????.xls");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            hworkbook.write(fos);
            
			System.out.println();
			System.out.println("--------------- ???? ???? ???????? ????. ---------------");
			System.out.println();
			
        } catch (FileNotFoundException e) {
            System.out.println("???? : student ???? ?????? ???? ???? ????????????.");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(hworkbook!=null) hworkbook.close();
                if(fos!=null) fos.close();
                
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
		
		
	}

public static int print_simple(int count,int penalty,HSSFSheet sheet,HSSFRow row,HSSFCell cell,CellStyle TitleStyle,CellStyle HeadStyle,CellStyle BodyStyle,Student[] student) {
	
	row = sheet.createRow(count);
    cell=row.createCell(0);
    if(penalty>=3) {
    	cell.setCellValue("???? 3?????? ??????");
    }
    else {
    	cell.setCellValue("???? 3?????? ??????");	
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    cell=row.createCell(0);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(1);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(2);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    count++;
    
    int a = 0,b=0;
    
    if(penalty==3) {
    	a=3;
    	b=20;    	
    }
    else if(penalty==1) {
    	a=1;
    	b=3;
    }
    
    for(int i=0;i<student.length;i++) {
    	for(int j=0;j<student[i].date.size();j++) {
    		if(student[i].penalty>=a&&student[i].penalty<b) { // if???? ???? ???????? ?????????? ?? ???????????????
    			row=sheet.createRow(count);
    			
    			cell = row.createCell(0);
    			cell.setCellValue(student[i].number.get(j));
    	        cell.setCellStyle(BodyStyle);
    			
                cell = row.createCell(1);
    			cell.setCellValue(student[i].username.get(j));
    	        cell.setCellStyle(BodyStyle);
    	        
    	        cell=row.createCell(2);
    			String s="";
    			for(int k=0;k<student[i].because.size();k++) {
    				s+=student[i].because.get(k)+" ";
    			}
    			cell.setCellValue(s);
    	        cell.setCellStyle(BodyStyle);
    			
    			cell = row.createCell(3);
    			cell.setCellValue(student[i].penalty+" ??");
    	        cell.setCellStyle(BodyStyle);
    			
    			count++;
    			break;
    		}
    		
    	}
    }
	
	return count;
}

public static int print_detail(int count,int penalty,HSSFSheet sheet,HSSFRow row,HSSFCell cell,CellStyle TitleStyle,CellStyle HeadStyle,CellStyle BodyStyle,Student[] student) {
	
	row=sheet.createRow(count);
    
    cell=row.createCell(0);
    if(penalty>=3) {
    	cell.setCellValue("???? 3?????? ?????? ????????");
    }
    else {
    	cell.setCellValue("???? 3?????? ?????? ????????");
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    
    // ???? ???? ????
    cell=row.createCell(0);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);

    
    cell=row.createCell(1);
    cell.setCellValue("???? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(2);
    cell.setCellValue("???? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("???? ??????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(4);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(5);
    cell.setCellValue("?????? ??");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(6);
    cell.setCellValue("???? ????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(7);
    cell.setCellValue("???? ????");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(8);
    cell.setCellValue("????");
    cell.setCellStyle(HeadStyle);
    count++;
    
    int a=0,b=0;
    
    if(penalty==3) {
    	a=3;
    	b=20;
    }
    else if(penalty==1) {
    	a=1;
    	b=3;
    }
    

    Student vo;
    for(int rowIdx=0; rowIdx < student.length; rowIdx++) {
        vo = student[rowIdx];
    	for(int j=0;j<vo.tag.size();j++) {
    		int k=vo.tag.get(j);
    		
            if(student[rowIdx].penalty>=a&&student[rowIdx].penalty<b) {
            	
        	// ?? ????
            row = sheet.createRow(count++);
            
            cell = row.createCell(0);
            cell.setCellValue(vo.origin_date.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(1);
            cell.setCellValue(vo.origin_machine.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(2);
            cell.setCellValue(vo.origin_console.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(3);
            cell.setCellValue(vo.card.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(4);
            cell.setCellValue(vo.number.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(5);
            cell.setCellValue(vo.username.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(6);
            cell.setCellValue(vo.company.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(7);
            cell.setCellValue(vo.dept.get(k));
	        cell.setCellStyle(BodyStyle);
            
            cell = row.createCell(8);
            cell.setCellValue(vo.etc.get(k));
	        cell.setCellStyle(BodyStyle);
            
            }
    	}
    }
	
	return count;
}

}


public class Test {
	public static void pause() {
		try {
			System.out.println("Enter ???? ?????? ?????????? ??????????.");
			System.in.read();
		}catch(IOException e) {
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		
		// ???? ???????? ????
		JFileChooser chooser = new JFileChooser();
		JOptionPane.showMessageDialog(null,"?????? ?????? ????????????.");
				
		int ret = chooser.showOpenDialog(null);
		
		if(ret!=JFileChooser.APPROVE_OPTION) {
			JOptionPane.showMessageDialog(null,"?????? ???????? ????????????.","????",JOptionPane.WARNING_MESSAGE);
			return ;
		}
		
		// ???? ???? ????
		String filepath=chooser.getSelectedFile().getPath();
		
		int num=0;
		String str=filepath;
		
		while(num<filepath.length()) { // ?????????? ?????? ???? ?????? - ???? ?????????????? ??????
			num++;
			str=filepath.substring(filepath.length()-num,filepath.length());
			if(str.substring(0,1).equals("\\")) {
				break;
			}
		}
		
		// ????????
		String file_save_path=filepath.substring(0,filepath.length()-num+1);
		
		Scanner sc = new Scanner(System.in);
		System.out.println("?????? ?????? ?????? ??????.");
		System.out.println("ex) 2018?? 8?? 30?? ~ 31?????? '????????', '????????' ?? ?????? ???? : '20180831' ????");
		int date=sc.nextInt();
		
		String human[][];
		
		//???????? ???? ???? - which?????? xlsx, xls?????? ???????? ???? ?????? ????????????
		String which=filepath.substring(filepath.length()-3, filepath.length());

		if(which.equals("lsx")) {
			human=xlsx.read(filepath);
		}
		else {
			human=xls.read(filepath);
		}

		int row=human.length;

		/* readExcel() ???????? ???? human?? ???? ?? ???????? ???????? ????
		for(int i=0;i<row;i++) {
          	for(int j=0;j<human[i].length;j++) {
            	System.out.print(human[i][j]+" ");
          	}
           	System.out.println();
        }*/
		
			//dom_total???? ???? ?????? ???????? ?????????? ???????? ex)000046,000716 ??..
            ArrayList<String> dom_total=new ArrayList<>(); // dom_total?? ???????? ?? ?????????? ?????????? ??
			for(int i=1;i<row;i++) { // ?????? ?? ???? ???? ???? - 1???? ???????? ?????? ?????? ????,?????? ?????? ?????? ???? ?? 
				boolean flag=false;   
				for(String value2:dom_total) {  // dom_total?? ???????? ?????? ?????? ???????? ?????? flag?? true ????
					if(value2.equals(human[i][3])) {
						flag=true;
						break;
						}
					}
				if(!flag) {						// ?????? ??????. ??, flag false???? ????
				dom_total.add(human[i][3]);
				}
			}
			
			// dom_total_size???? ???? ?????? ?????? ???? ????????
			int dom_total_size=dom_total.size();
			
			Student[] student=new Student[dom_total_size]; // ???? ?? ???? student ???? ???? ????
			for(int i=0;i<dom_total_size;i++) {            // ???? ?? ???? ???? 
				
				student[i]=new Student();                  // ???? ?? ???? ???? ????
				
				//?????? ???? ?????? ?????? ???????? ?????? ????????
				for(int j=0;j<row;j++) { // human[][]?? ?? ???????? ?????????? ?????? ????
					if(dom_total.get(i).equals(human[j][3])) {
						student[i].card.add(dom_total.get(i));
						student[i].origin_date.add(human[j][0]);
						String date_time=human[j][0]; // human ?????? ???? ?????? student ?????? ????, ?????? ?????? ???????? ????
						String result1=date_time.substring(0,11); 
						String intresult1=result1.replaceAll("[^0-9]","");
						String result2=date_time.substring(date_time.lastIndexOf(" ")+1);
						String intresult2=result2.replaceAll("[^0-9]","");
						int hours=Integer.parseInt(intresult2.substring(0,2)); // ?????? ???????? ?????? time?????? ???????? ????
						int minutes=Integer.parseInt(intresult2.substring(2,4));
						int seconds=Integer.parseInt(intresult2.substring(4,6));
						int time=hours*3600+minutes*60+seconds;
						
						student[i].date.add(Integer.parseInt(intresult1));
						student[i].origin_time.add(Integer.parseInt(intresult2));
						student[i].time.add(time);
						student[i].origin_machine.add(human[j][1]);
						String machine=human[j][1];
						if(machine.indexOf("LANE")==1){
							student[i].machine.add("LANE");
						}
						else {
							student[i].machine.add("NO LANE");
						}
												
						student[i].origin_console.add(human[j][2]);
						
						String console=human[j][2];  // '1??????-IN'?? ?????? ?????????? IN OUT?? ???????? ????
						if(console.indexOf("IN")==-1) {
							student[i].console.add("OUT");
						}
						else {
							student[i].console.add("IN");
						}
						
						student[i].number.add(human[j][4]);
						student[i].username.add(human[j][5]);
						student[i].company.add(human[j][6]);
						student[i].dept.add(human[j][7]);
						student[i].etc.add(human[j][8]);
					}
				}
			}
			
			//?????? ???????? ???????? ?????? ???????????? ???????? ????
			/*
			for(int i=0;i<dom_total.size();i++) {
				for(int j=0;j<student[i].date.size();j++) {
					System.out.print(student[i].date.get(j)+" ");
					System.out.print(student[i].origin_time.get(j)+" ");
					System.out.print(student[i].origin_machine.get(j)+" ");
					System.out.print(student[i].console.get(j)+" ");
					System.out.print(student[i].card.get(j)+" ");
					System.out.print(student[i].number.get(j)+" ");
					System.out.print(student[i].username.get(j)+" ");
					System.out.print(student[i].company.get(j)+" ");
					System.out.print(student[i].dept.get(j)+" ");
					System.out.print(student[i].etc.get(j)+" ");
					System.out.println();
				}
				System.out.println();
			}
			*/
			
			// ?????????? ????
			for(int i=0;i<student.length;i++) { // ???? student ????
				for(int j=0;j<student[i].date.size();j++) { // student 1?? ???? ????
					if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="OUT"&&student[i].machine.get(j)=="LANE") {
						// ????????
						
						boolean cigarette_time=false; //???? false
						
						for(int k=j+1;k<=student[i].date.size();k++) { //?????? ???????? 10?? ????
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+600>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								cigarette_time=true;
								j=k;
							}
						}
						if(cigarette_time) { //?????????? ????
							//System.out.println(student[i].username.get(j)+" ????");
						}
						else{ // ???? ???????? ????????
							for(int k=j+1;k<=student[i].date.size();k++) { //?????? ???????? 20???? ???????????? ?????? 1???? ??????????
								if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="OUT"&&student[i].machine.get(k)=="LANE") {
									j=k;
								}
							}
							//System.out.println(student[i].username.get(j)+" ????????");
							student[i].because.add("????????");
							student[i].penalty+=2;

						}
					}
					else if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="IN"&&student[i].machine.get(j)=="LANE") {
						// ????????
						
						for(int k=j+1;k<=student[i].date.size();k++) { //?????? ???????? 20???? ???????????? ?????? 1???? ??????????
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								j=k;
							}
						}
						//System.out.println(student[i].username.get(j)+" ????????");
						student[i].because.add("????????");
						student[i].penalty++;
						
					}
				}
				
				if(student[i].penalty>=3) {
					student[i].caution=true;
				}
				
				for(int j=0;j<student[i].date.size();j++) {
					if(student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000) {
						student[i].tag.add(j);
					}
				}
			}
			
			  // ?????? ???? ?????? ????
			//xlsx.write(student,file_save_path,date);
			
			if(which.equals("lsx")) {
				xlsx.write(student,file_save_path,date);
			}
			else {
				xls.write(student,file_save_path,date);
			}

			// ?????? ???? ????
			System.out.println();
			System.out.println("---------------------???? ????---------------------");
			System.out.println("???? ???????? ???? ?????????? ?????????? ?????? ?????? ????????????.");
			System.out.println();
			for(int i=0;i<student.length;i++) {
				if(student[i].caution==true) { // ???? 3?????????? ????
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"?? ???? : "+student[i].penalty+" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println();
			System.out.println("-------------------???? ???? ????-------------------");
			System.out.println("???? ???????? ???? ???? ?????????? ??????.");
			System.out.println();
			for(int i=0;i<student.length;i++) {
				if(student[i].penalty>=1&&student[i].penalty<3) { // ???? 3?? ???????? ????
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"?? ???? : "+student[i].penalty +" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println("???????? ?????? ??????????????.");
			System.out.println();
			pause();
			
	}
	
}
