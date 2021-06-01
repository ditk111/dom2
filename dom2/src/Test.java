import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.text.html.HTMLDocument.Iterator;

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

class Excel{			
public static String[][] readExcel() {
		try {
			FileInputStream file = new FileInputStream("C:\\Users\\user\\Desktop\\개발관련자료\\dom2\\2.xlsx");
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	           
	        XSSFSheet sheet=workbook.getSheetAt(0);   //시트 수 (첫번째에만 존재하므로 0을 준다) / 만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
	        String value="";   // 담을 값 초기화
	        int rowindex=0;    // 행 인덱스 초기화
	        int columnindex=0; // 열 인덱스 초기화
	        int rows=sheet.getPhysicalNumberOfRows();// 총 행의 수
	            
	        XSSFRow firstrow=sheet.getRow(rowindex); //첫행 읽어오기
	        int firstcells=firstrow.getPhysicalNumberOfCells(); // 첫행의 열의 수 구하기
	            
	        String[][] human=new String[rows][firstcells]; //2차원배열 초기화

	        for(rowindex=0;rowindex<rows;rowindex++){
	            XSSFRow row=sheet.getRow(rowindex); //행 읽어오기
	            if(row !=null){
	                int cells=row.getPhysicalNumberOfCells(); // 열의 수 읽어오기
	                for(columnindex=0; columnindex<=cells; columnindex++){
	                    //셀값을 읽는다
	                    XSSFCell cell=row.getCell(columnindex);
	                    /*String value="";
	                    //셀이 빈값일경우를 위한 널체크*/
	                    if(cell==null){
	                        continue;
	                    }else{
	                        //타입별로 내용 읽기
	                        switch (cell.getCellType()){
	                        case FORMULA:
	                            value=cell.getCellFormula();
	                            break;
	                        case NUMERIC:
	                            value=cell.getNumericCellValue()+"";
	                            break;
	                        case STRING: // 실제로는 String 값만 받아들임. 빈칸도 String으로 받아들임.
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

public static void writeExcel(Student[] student) {
	
		 // 워크북 생성
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        // 워크시트 생성
        XSSFSheet sheet = workbook.createSheet();
        // 행 생성
        XSSFRow row=sheet.createRow(1);
        // 쎌 생성
        XSSFCell cell=row.createCell(0);
        
        //title
        XSSFFont TitleFont=workbook.createFont();
        TitleFont.setFontHeightInPoints((short)13);
        TitleFont.setFontName("맑은 고딕");
        TitleFont.setColor(IndexedColors.BLUE.getIndex());
        TitleFont.setBold(true);
        
        CellStyle TitleStyle=workbook.createCellStyle();
        TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        TitleStyle.setFont(TitleFont);
        
        //Head
        XSSFFont HeadFont=workbook.createFont();
        HeadFont.setBold(true);
        HeadFont.setFontHeightInPoints((short)13);
        HeadFont.setFontName("맑은 고딕");
        
        CellStyle HeadStyle=workbook.createCellStyle();
        HeadStyle.setAlignment(HorizontalAlignment.CENTER);
        HeadStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        HeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HeadStyle.setBorderTop(BorderStyle.THIN);
        HeadStyle.setBorderBottom(BorderStyle.THIN);
        HeadStyle.setBorderLeft(BorderStyle.THIN);
        HeadStyle.setBorderRight(BorderStyle.THIN);
        HeadStyle.setFont(HeadFont);
        
        //Body
        XSSFFont BodyFont=workbook.createFont();
        BodyFont.setFontHeightInPoints((short)11);
        BodyFont.setFontName("맑은 고딕");
        
        CellStyle BodyStyle=workbook.createCellStyle();
        BodyStyle.setAlignment(HorizontalAlignment.CENTER);
        BodyStyle.setBorderTop(BorderStyle.THIN);
        BodyStyle.setBorderBottom(BorderStyle.THIN);
        BodyStyle.setBorderLeft(BorderStyle.THIN);
        BodyStyle.setBorderRight(BorderStyle.THIN);
        BodyStyle.setFont(BodyFont);
        
        // 여백
        sheet.setColumnWidth(0, (short)8000); //일시
        sheet.setColumnWidth(1, (short)3500); //기기명
        sheet.setColumnWidth(2, (short)6500); //콘솔명
        sheet.setColumnWidth(3, (short)3000); //카드 아이디
        sheet.setColumnWidth(4, (short)3000); // 사번
        sheet.setColumnWidth(5, (short)3500); // 사용자 명
        sheet.setColumnWidth(6, (short)3000); // 근무회사
        sheet.setColumnWidth(7, (short)3000); // 근무부서
        sheet.setColumnWidth(8, (short)3000); // 구분

        int count=1; // count는 행 위치를 나타냄
        

        count=print_simple(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// 벌점 3점이상 사생들

        count++;
        count=print_detail(count,3,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// 벌점 3점이상 사생들 상세목록
        
        count++;
        count=print_simple(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// 벌점 3점미만 사생들
                
        count++;
        count=print_detail(count,1,sheet,row,cell,TitleStyle,HeadStyle,BodyStyle,student);	// 벌점 3점미만 사생들 상세목록
        
        // 입력된 내용 파일로 쓰기
        File file = new File("C:\\Users\\user\\Desktop\\개발관련자료\\dom2\\student.xlsx");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
            
			System.out.println();
			System.out.println("--------------- 엑셀 파일 다운로드 완료. ---------------");
			System.out.println();
			
        } catch (FileNotFoundException e) {
            System.out.println("에러 : student 엑셀 파일을 닫고 다시 실행해주세요.");
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

public static int print_simple(int count,int penalty,XSSFSheet sheet,XSSFRow row,XSSFCell cell,CellStyle TitleStyle,CellStyle HeadStyle,CellStyle BodyStyle,Student[] student) {
	
	row = sheet.createRow(count);
    cell=row.createCell(0);
    if(penalty>=3) {
    	cell.setCellValue("벌점 3점이상 사생들");
    }
    else {
    	cell.setCellValue("벌점 3점미만 사생들");	
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    cell=row.createCell(0);
    cell.setCellValue("사번");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(1);
    cell.setCellValue("이름");
    cell.setCellStyle(HeadStyle);

    cell=row.createCell(2);
    cell.setCellValue("사유");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("벌점");
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
    		if(student[i].penalty>=a&&student[i].penalty<b) { // if문을 한칸 밖에하면 조금이라도 더 빨라지지않을까?
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
    			cell.setCellValue(student[i].penalty+" 점");
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
    	cell.setCellValue("벌점 3점이상 사생들 상세목록");
    }
    else {
    	cell.setCellValue("벌점 3점미만 사생들 상세목록");
    }
    cell.setCellStyle(TitleStyle);
    count++;
    
    row=sheet.createRow(count);
    
    // 헤더 정보 구성
    cell=row.createCell(0);
    cell.setCellValue("일시");
    cell.setCellStyle(HeadStyle);

    
    cell=row.createCell(1);
    cell.setCellValue("기기 명");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(2);
    cell.setCellValue("콘솔 명");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(3);
    cell.setCellValue("카드 아이디");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(4);
    cell.setCellValue("사번");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(5);
    cell.setCellValue("사용자 명");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(6);
    cell.setCellValue("근무 회사");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(7);
    cell.setCellValue("근무 부서");
    cell.setCellStyle(HeadStyle);
    
    cell=row.createCell(8);
    cell.setCellValue("구분");
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
            	
        	// 행 생성
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
			System.out.println("Enter 키를 입력시 프로그램이 종료됩니다.");
			System.in.read();
		}catch(IOException e) {
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		
		Scanner sc = new Scanner(System.in);
		System.out.println("검사할 날짜를 입력해 주세요.");
		System.out.println("ex) 2018년 8월 30일 ~ 31일간의 '지연복귀', '심야이탈' 을 검사할 경우 : '20180831' 입력");
		int date=sc.nextInt();
		
		String human[][];
		//엑셀파일 읽는 부분
		human=Excel.readExcel();
		
		int row=human.length;

		/* readExcel() 메소드를 통해 human에 값이 잘 실렸는지 확인하는 구문
		for(int i=0;i<row;i++) {
          	for(int j=0;j<human[i].length;j++) {
            	System.out.print(human[i][j]+" ");
          	}
           	System.out.println();
        }*/
		
			//dom_total에는 금일 출입한 학생들의 카드번호가 들어있음 ex)000046,000716 등..
            ArrayList<String> dom_total=new ArrayList<>(); // dom_total에 중복값을 뺀 카드번호만 삽입하고자 함
			for(int i=1;i<row;i++) { // 엑셀의 총 행수 만큼 돌기 - 1부터 시작하는 이유는 첫행의 일시,기기명 이런걸 빼주기 위한 것 
				boolean flag=false;   
				for(String value2:dom_total) {  // dom_total에 중복값이 있는지 확인후 중복값이 있으면 flag에 true 주기
					if(value2.equals(human[i][3])) {
						flag=true;
						break;
						}
					}
				if(!flag) {						// 중복값 없을때. 즉, flag false일때 삽입
				dom_total.add(human[i][3]);
				}
			}
			
			// dom_total_size에는 금일 출입한 학생들 수가 들어있음
			int dom_total_size=dom_total.size();
			
			Student[] student=new Student[dom_total_size]; // 학생 수 만큼 student 배열 길이 선언
			for(int i=0;i<dom_total_size;i++) {            // 학생 수 만큼 반복 
				
				student[i]=new Student();                  // 학생 수 만큼 객체 생성
				
				//반복문 계속 돌면서 카드가 일치하면 데이터 삽입구문
				for(int j=0;j<row;j++) { // human[][]의 행 길이만큼 반복하면서 데이터 삽입
					if(dom_total.get(i).equals(human[j][3])) {
						student[i].card.add(dom_total.get(i));
						student[i].origin_date.add(human[j][0]);
						String date_time=human[j][0]; // human 배열에 있는 날짜를 student 객체의 날짜, 시간을 나눠서 숫자값만 넣기
						String result1=date_time.substring(0,11); 
						String intresult1=result1.replaceAll("[^0-9]","");
						String result2=date_time.substring(date_time.lastIndexOf(" ")+1);
						String intresult2=result2.replaceAll("[^0-9]","");
						int hours=Integer.parseInt(intresult2.substring(0,2)); // 시간을 초단위로 나눠서 time변수에 대입하는 과정
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
						
						String console=human[j][2];  // '1층정문-IN'와 유사한 데이터에서 IN OUT만 식별하여 삽입
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
			
			//반복문 돌리면서 데이터가 제대로 삽입되었는지 확인하는 구문
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
			
			// 벌점메기는 구문
			for(int i=0;i<student.length;i++) { // 모든 student 반복
				for(int j=0;j<student[i].date.size();j++) { // student 1명 내부 반복
					if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="OUT"&&student[i].machine.get(j)=="LANE") {
						// 심야이탈
						
						boolean cigarette_time=false; //담타 false
						
						for(int k=j+1;k<=student[i].date.size();k++) { //반복문 담배타임 10분 체크
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+600>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								cigarette_time=true;
								j=k;
							}
						}
						if(cigarette_time) { //담타일경우 담타
							//System.out.println(student[i].username.get(j)+" 담타");
						}
						else{ // 담타 아닐경우 심야이탈
							for(int k=j+1;k<=student[i].date.size();k++) { //반복문 지연복귀 20초간 기기에러인지 확인후 1번만 출력되도록
								if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="OUT"&&student[i].machine.get(k)=="LANE") {
									j=k;
								}
							}
							//System.out.println(student[i].username.get(j)+" 심야이탈");
							student[i].because.add("심야이탈");
							student[i].penalty+=2;

						}
					}
					else if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="IN"&&student[i].machine.get(j)=="LANE") {
						// 지연복귀
						
						for(int k=j+1;k<=student[i].date.size();k++) { //반복문 지연복귀 20초간 기기에러인지 확인후 1번만 출력되도록
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								j=k;
							}
						}
						//System.out.println(student[i].username.get(j)+" 지연복귀");
						student[i].because.add("지연복귀");
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
			
			  // 입력된 내용 파일로 쓰기
			Excel.writeExcel(student);

			// 데이터 출력 구문
			System.out.println();
			System.out.println("---------------------주의 사생---------------------");
			System.out.println("아래 사생들은 새로 다운로드된 엑셀파일을 확인후 벌점을 부여해주세요.");
			System.out.println();
			for(int i=0;i<student.length;i++) {
				if(student[i].caution==true) { // 벌점 3점이상자들 출력
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"의 벌점 : "+student[i].penalty+" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println();
			System.out.println("-------------------금일 벌점 사생-------------------");
			System.out.println("아래 사생들은 바로 벌점 부여하셔도 됩니다.");
			System.out.println();
			for(int i=0;i<student.length;i++) {
				if(student[i].penalty>=1&&student[i].penalty<3) { // 벌점 3점 미만자들 출력
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"의 벌점 : "+student[i].penalty +" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println("프로그램 실행이 완료되었습니다.");
			System.out.println();
			pause();
			
	}
	
}
