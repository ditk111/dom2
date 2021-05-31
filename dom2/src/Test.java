import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
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
	ArrayList<String> tag=new ArrayList<>();
	int penalty=0;
	boolean caution=false;
}

public class Test {
	


	public static void main(String[] args) {
		try {
			Scanner sc = new Scanner(System.in);
			System.out.println("검사할 날짜를 입력해 주세요.");
			System.out.println("ex) 2018년 8월 30일 ~ 31일간의 무단외출,외박을 검사할 경우 : '20180831' 입력");
			int date=sc.nextInt();
			
			
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
            /*for(int i=0;i<rows;i++) {
            	for(int j=0;j<firstcells;j++) {
            		System.out.print(human[i][j]+" ");
            	}
            	System.out.println();
            }*/
            
            ArrayList<String> dom_total=new ArrayList<>(); // dom_total에 중복값을 뺀 카드번호만 삽입하고자 함
			for(int i=1;i<rows;i++) { // 엑셀의 총 행수 만큼 돌기 / 1만큼 시작하는 이유는 첫행의 일시,기기명 이런걸 빼주기 위한 것 
				boolean flag=false;   
				for(String value2:dom_total) {  // dom_total에 중복값이 있는지 확인후 중복값이 없으면 flag에 true 주기
					if(value2.equals(human[i][3])) {
						flag=true;
						break;
						}
					}
				if(!flag) {						// falg true일때 데이터 삽입
				dom_total.add(human[i][3]);
				}
			}
			//System.out.println(dom_total);		// 총 학생들의 카드 넘버
			//System.out.println(dom_total.size()); // 총 카드넘버 수
			
			Student[] student=new Student[dom_total.size()]; // 학생 수 만큼 student 배열 길이 선언
			for(int i=0;i<dom_total.size();i++) {            // 학생 수 만큼 반복 
				
				student[i]=new Student();                    // 학생 수 만큼 객체 생성
				
				//반복문 계속 돌면서 카드가 일치하면 데이터 삽입구현
				for(int j=0;j<rows;j++) { // j는 행을 가리킴 / 반복문 돌면서 객체별로(학생별로) 데이터 삽입
					if(dom_total.get(i).equals(human[j][3])) {
						student[i].card.add(dom_total.get(i));
						student[i].origin_date.add(human[j][0]);
						String date_time=human[j][0]; // human 배열에 있는 날짜를 student 클래스의 날짜, 시간을 나눠서 숫자값만 넣기
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
			
			//반복문 돌리면서 데이터출력 체크
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
			

			for(int i=0;i<dom_total.size();i++) { // 기숙사생들 일일이 반복문돌면서 비교
				for(int j=0;j<student[i].date.size();j++) { // 기숙사생 1명씩 사생들의 데이터 비교
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
			}
			System.out.println();
			System.out.println("엑셀 파일 다운로드 완료.");
			  // 입력된 내용 파일로 쓰기
			writeExcel(student);


			
			System.out.println();
			System.out.println("---------------------주의 사생---------------------");
			System.out.println("아래 사생들은 새로 다운로드된 엑셀파일을 확인후 벌점을 부여해주세요.");
			System.out.println();
			for(int i=0;i<dom_total.size();i++) { // 기숙사생들 일일이 반복문돌면서 비교
				if(student[i].caution==true) {
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"의 벌점 : "+student[i].penalty+" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println();
			System.out.println("-------------------금일 벌점 사생-------------------");
			System.out.println("아래 사생들은 바로 벌점 부여하셔도 됩니다.");
			System.out.println();
			for(int i=0;i<dom_total.size();i++) { // 기숙사생들 일일이 반복문돌면서 비교
				if(student[i].penalty!=0&&student[i].penalty<3) {
					System.out.println( student[i].number.get(0)+" "+student[i].username.get(0)+"의 벌점 : "+student[i].penalty +" "+student[i].because);
				}
			}
			

			

			
			/* 내일 할것 ! 
			 * 클래스에 caution 변수 불린으로 선언후 벌점 많은사람들만 별도로 표시해야함. - 완료
			 * 데이터가 잘 출력되는지 검토해야함 - 이상없음 확인완료.
			 * caution과 일반벌점자들 엑셀파일로 삽입하는 작업
		*/
			
			/* 다음에 할것 !
			 * 리팩토링 !!
			 */
			
 
        }catch(Exception e) {
            e.printStackTrace();
        }
	}
	
	
	
	public static void writeExcel(Student[] student) {
		 // 워크북 생성
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 워크시트 생성
        XSSFSheet sheet = workbook.createSheet();
        // 행 생성
        XSSFRow row = sheet.createRow(0);
        // 쎌 생성
        XSSFCell cell;
        
        int count=0;
        

        cell=row.createCell(4);
        count++;
        cell.setCellValue("벌점 3점이상 사생들");
        
        row=sheet.createRow(count);
        count++;
        cell=row.createCell(4);
        cell.setCellValue("사번");
        
        cell=row.createCell(5);
        cell.setCellValue("이름");
        
        cell=row.createCell(6);
        cell.setCellValue("벌점");
        
        cell=row.createCell(7);
        cell.setCellValue("사유");

        
        for(int i=0;i<student.length;i++) {
        	for(int j=0;j<student[i].date.size();j++) {
        		if(student[i].penalty>=3) {
        			row=sheet.createRow(count);
        			
        			cell = row.createCell(4);
        			cell.setCellValue(student[i].number.get(j));
        			
                    cell = row.createCell(5);
        			cell.setCellValue(student[i].username.get(j));
        			
        			cell = row.createCell(6);
        			cell.setCellValue(student[i].penalty+" 점");
        			
        			cell=row.createCell(7);
        			String s="";
        			for(int k=0;k<student[i].because.size();k++) {
        				s+=student[i].because.get(k)+" ";
        			}
        			cell.setCellValue(s);
        			count++;
        			break;
        		}
        	}
        }
        
        count++;
        row=sheet.createRow(count);
        count++;
        
        cell=row.createCell(0);
        cell.setCellValue("벌점 3점이상 사생들 상세목록");
        
        row=sheet.createRow(count);
        count++;
        
        // 헤더 정보 구성
        cell=row.createCell(0);
        cell.setCellValue("일시");
        
        cell=row.createCell(1);
        cell.setCellValue("기기 명");
        
        cell=row.createCell(2);
        cell.setCellValue("콘솔 명");
        
        cell=row.createCell(3);
        cell.setCellValue("카드 아이디");
        
        cell=row.createCell(4);
        cell.setCellValue("사번");
        
        cell=row.createCell(5);
        cell.setCellValue("사용자 명");
        
        cell=row.createCell(6);
        cell.setCellValue("근무 회사");
        
        cell=row.createCell(7);
        cell.setCellValue("근무 부서");
        
        cell=row.createCell(8);
        cell.setCellValue("구분");
        

        
        // 리스트의 size 만큼 row를 생성
        
        //3점이상들은 1번만 출력후 아래 상세표시되게 해야함
        Student vo;
        for(int rowIdx=0; rowIdx < student.length; rowIdx++) {

        	for(int j=0;j<student[rowIdx].date.size();j++) {
        		
            vo = student[rowIdx];
	            if(vo.penalty>=3) {

	            
	        	// 행 생성
	            row = sheet.createRow(count++);
	            
	            
	            cell = row.createCell(0);
	            cell.setCellValue(vo.origin_date.get(j));
	            
	            cell = row.createCell(1);
	            cell.setCellValue(vo.origin_machine.get(j));
	            
	            cell = row.createCell(2);
	            cell.setCellValue(vo.origin_console.get(j));
	            
	            cell = row.createCell(3);
	            cell.setCellValue(vo.card.get(j));
	            
	            cell = row.createCell(4);
	            cell.setCellValue(vo.number.get(j));
	            
	            cell = row.createCell(5);
	            cell.setCellValue(vo.username.get(j));
	            
	            cell = row.createCell(6);
	            cell.setCellValue(vo.company.get(j));
	            
	            cell = row.createCell(7);
	            cell.setCellValue(vo.dept.get(j));
	            
	            cell = row.createCell(8);
	            cell.setCellValue(vo.etc.get(j));
	            
	            
	            /*
	            cell = row.createCell(3);
	            cell.setCellValue(vo.penalty);*/
	            }
        	}
        }
        
        count++;
        row = sheet.createRow(count);
        cell=row.createCell(4);
        count++;
        cell.setCellValue("벌점 3점미만 사생들");
        
        row=sheet.createRow(count);
        count++;
        cell=row.createCell(4);
        cell.setCellValue("사번");
        
        cell=row.createCell(5);
        cell.setCellValue("이름");
        
        cell=row.createCell(6);
        cell.setCellValue("벌점");
        
        cell=row.createCell(7);
        cell.setCellValue("사유");

        
        for(int i=0;i<student.length;i++) {
        	for(int j=0;j<student[i].date.size();j++) {
        		if(student[i].penalty>0&&student[i].penalty<3) {
        			row=sheet.createRow(count);
        			
        			cell = row.createCell(4);
        			cell.setCellValue(student[i].number.get(j));
        			
                    cell = row.createCell(5);
        			cell.setCellValue(student[i].username.get(j));
        			
        			cell = row.createCell(6);
        			cell.setCellValue(student[i].penalty+" 점");
        			
        			cell=row.createCell(7);
        			String s="";
        			for(int k=0;k<student[i].because.size();k++) {
        				s+=student[i].because.get(k)+" ";
        			}
        			cell.setCellValue(s);
        			count++;
        			break;
        		}
        	}
        }
        
        
        
        // 입력된 내용 파일로 쓰기
        File file = new File("C:\\Users\\user\\Desktop\\개발관련자료\\dom2\\student.xlsx");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
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
}
