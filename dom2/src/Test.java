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
			System.out.println("�˻��� ��¥�� �Է��� �ּ���.");
			System.out.println("ex) 2018�� 8�� 30�� ~ 31�ϰ��� ���ܿ���,�ܹ��� �˻��� ��� : '20180831' �Է�");
			int date=sc.nextInt();
			
			
            FileInputStream file = new FileInputStream("C:\\Users\\user\\Desktop\\���߰����ڷ�\\dom2\\2.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            
            XSSFSheet sheet=workbook.getSheetAt(0);   //��Ʈ �� (ù��°���� �����ϹǷ� 0�� �ش�) / ���� �� ��Ʈ�� �б����ؼ��� FOR���� �ѹ��� �����ش�
            String value="";   // ���� �� �ʱ�ȭ
            int rowindex=0;    // �� �ε��� �ʱ�ȭ
            int columnindex=0; // �� �ε��� �ʱ�ȭ
            int rows=sheet.getPhysicalNumberOfRows();// �� ���� ��
            
            XSSFRow firstrow=sheet.getRow(rowindex); //ù�� �о����
            int firstcells=firstrow.getPhysicalNumberOfCells(); // ù���� ���� �� ���ϱ�
            
            String[][] human=new String[rows][firstcells]; //2�����迭 �ʱ�ȭ

            for(rowindex=0;rowindex<rows;rowindex++){
                XSSFRow row=sheet.getRow(rowindex); //�� �о����
                if(row !=null){
                    int cells=row.getPhysicalNumberOfCells(); // ���� �� �о����
                    for(columnindex=0; columnindex<=cells; columnindex++){
                        //������ �д´�
                        XSSFCell cell=row.getCell(columnindex);
                        /*String value="";
                        //���� ���ϰ�츦 ���� ��üũ*/
                        if(cell==null){
                            continue;
                        }else{
                            //Ÿ�Ժ��� ���� �б�
                            switch (cell.getCellType()){
                            case FORMULA:
                                value=cell.getCellFormula();
                                break;
                            case NUMERIC:
                                value=cell.getNumericCellValue()+"";
                                break;
                            case STRING: // �����δ� String ���� �޾Ƶ���. ��ĭ�� String���� �޾Ƶ���.
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
            
            ArrayList<String> dom_total=new ArrayList<>(); // dom_total�� �ߺ����� �� ī���ȣ�� �����ϰ��� ��
			for(int i=1;i<rows;i++) { // ������ �� ��� ��ŭ ���� / 1��ŭ �����ϴ� ������ ù���� �Ͻ�,���� �̷��� ���ֱ� ���� �� 
				boolean flag=false;   
				for(String value2:dom_total) {  // dom_total�� �ߺ����� �ִ��� Ȯ���� �ߺ����� ������ flag�� true �ֱ�
					if(value2.equals(human[i][3])) {
						flag=true;
						break;
						}
					}
				if(!flag) {						// falg true�϶� ������ ����
				dom_total.add(human[i][3]);
				}
			}
			//System.out.println(dom_total);		// �� �л����� ī�� �ѹ�
			//System.out.println(dom_total.size()); // �� ī��ѹ� ��
			
			Student[] student=new Student[dom_total.size()]; // �л� �� ��ŭ student �迭 ���� ����
			for(int i=0;i<dom_total.size();i++) {            // �л� �� ��ŭ �ݺ� 
				
				student[i]=new Student();                    // �л� �� ��ŭ ��ü ����
				
				//�ݺ��� ��� ���鼭 ī�尡 ��ġ�ϸ� ������ ���Ա���
				for(int j=0;j<rows;j++) { // j�� ���� ����Ŵ / �ݺ��� ���鼭 ��ü����(�л�����) ������ ����
					if(dom_total.get(i).equals(human[j][3])) {
						student[i].card.add(dom_total.get(i));
						student[i].origin_date.add(human[j][0]);
						String date_time=human[j][0]; // human �迭�� �ִ� ��¥�� student Ŭ������ ��¥, �ð��� ������ ���ڰ��� �ֱ�
						String result1=date_time.substring(0,11); 
						String intresult1=result1.replaceAll("[^0-9]","");
						String result2=date_time.substring(date_time.lastIndexOf(" ")+1);
						String intresult2=result2.replaceAll("[^0-9]","");
						int hours=Integer.parseInt(intresult2.substring(0,2)); // �ð��� �ʴ����� ������ time������ �����ϴ� ����
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
						
						String console=human[j][2];  // '1������-IN'�� ������ �����Ϳ��� IN OUT�� �ĺ��Ͽ� ����
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
			
			//�ݺ��� �����鼭 ��������� üũ
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
			

			for(int i=0;i<dom_total.size();i++) { // �������� ������ �ݺ������鼭 ��
				for(int j=0;j<student[i].date.size();j++) { // ������ 1�� ������� ������ ��
					if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="OUT"&&student[i].machine.get(j)=="LANE") {
						// �ɾ���Ż
						
						boolean cigarette_time=false; //��Ÿ false
						
						for(int k=j+1;k<=student[i].date.size();k++) { //�ݺ��� ���Ÿ�� 10�� üũ
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+600>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								cigarette_time=true;
								j=k;
							}
						}
						if(cigarette_time) { //��Ÿ�ϰ�� ��Ÿ
							//System.out.println(student[i].username.get(j)+" ��Ÿ");
						}
						else{ // ��Ÿ �ƴҰ�� �ɾ���Ż
							for(int k=j+1;k<=student[i].date.size();k++) { //�ݺ��� �������� 20�ʰ� ��⿡������ Ȯ���� 1���� ��µǵ���
								if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="OUT"&&student[i].machine.get(k)=="LANE") {
									j=k;
								}
							}
							//System.out.println(student[i].username.get(j)+" �ɾ���Ż");
							student[i].because.add("�ɾ���Ż");
							student[i].penalty+=2;

						}
					}
					else if(student[i].date.get(j)==date&&student[i].time.get(j)>=3600&&student[i].time.get(j)<=18000&&student[i].console.get(j)=="IN"&&student[i].machine.get(j)=="LANE") {
						// ��������
						
						for(int k=j+1;k<=student[i].date.size();k++) { //�ݺ��� �������� 20�ʰ� ��⿡������ Ȯ���� 1���� ��µǵ���
							if(student[i].date.size()>=k+1&&student[i].date.get(k)==date&&student[i].time.get(k)>=3600&& student[i].time.get(j)+20>=student[i].time.get(k) &&student[i].console.get(k)=="IN"&&student[i].machine.get(k)=="LANE") {
								j=k;
							}
						}
						//System.out.println(student[i].username.get(j)+" ��������");
						student[i].because.add("��������");
						student[i].penalty++;

					}
				}
				
			if(student[i].penalty>=3) {
				student[i].caution=true;
			}			
			}
			System.out.println();
			System.out.println("���� ���� �ٿ�ε� �Ϸ�.");
			  // �Էµ� ���� ���Ϸ� ����
			writeExcel(student);


			
			System.out.println();
			System.out.println("---------------------���� ���---------------------");
			System.out.println("�Ʒ� ������� ���� �ٿ�ε�� ���������� Ȯ���� ������ �ο����ּ���.");
			System.out.println();
			for(int i=0;i<dom_total.size();i++) { // �������� ������ �ݺ������鼭 ��
				if(student[i].caution==true) {
					System.out.println(student[i].number.get(0)+" "+student[i].username.get(0)+"�� ���� : "+student[i].penalty+" "+student[i].because);
				}
			}
			System.out.println();
			System.out.println();
			System.out.println("-------------------���� ���� ���-------------------");
			System.out.println("�Ʒ� ������� �ٷ� ���� �ο��ϼŵ� �˴ϴ�.");
			System.out.println();
			for(int i=0;i<dom_total.size();i++) { // �������� ������ �ݺ������鼭 ��
				if(student[i].penalty!=0&&student[i].penalty<3) {
					System.out.println( student[i].number.get(0)+" "+student[i].username.get(0)+"�� ���� : "+student[i].penalty +" "+student[i].because);
				}
			}
			

			

			
			/* ���� �Ұ� ! 
			 * Ŭ������ caution ���� �Ҹ����� ������ ���� ��������鸸 ������ ǥ���ؾ���. - �Ϸ�
			 * �����Ͱ� �� ��µǴ��� �����ؾ��� - �̻���� Ȯ�οϷ�.
			 * caution�� �Ϲݹ����ڵ� �������Ϸ� �����ϴ� �۾�
		*/
			
			/* ������ �Ұ� !
			 * �����丵 !!
			 */
			
 
        }catch(Exception e) {
            e.printStackTrace();
        }
	}
	
	
	
	public static void writeExcel(Student[] student) {
		 // ��ũ�� ����
        XSSFWorkbook workbook = new XSSFWorkbook();
        // ��ũ��Ʈ ����
        XSSFSheet sheet = workbook.createSheet();
        // �� ����
        XSSFRow row = sheet.createRow(0);
        // �� ����
        XSSFCell cell;
        
        int count=0;
        

        cell=row.createCell(4);
        count++;
        cell.setCellValue("���� 3���̻� �����");
        
        row=sheet.createRow(count);
        count++;
        cell=row.createCell(4);
        cell.setCellValue("���");
        
        cell=row.createCell(5);
        cell.setCellValue("�̸�");
        
        cell=row.createCell(6);
        cell.setCellValue("����");
        
        cell=row.createCell(7);
        cell.setCellValue("����");

        
        for(int i=0;i<student.length;i++) {
        	for(int j=0;j<student[i].date.size();j++) {
        		if(student[i].penalty>=3) {
        			row=sheet.createRow(count);
        			
        			cell = row.createCell(4);
        			cell.setCellValue(student[i].number.get(j));
        			
                    cell = row.createCell(5);
        			cell.setCellValue(student[i].username.get(j));
        			
        			cell = row.createCell(6);
        			cell.setCellValue(student[i].penalty+" ��");
        			
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
        cell.setCellValue("���� 3���̻� ����� �󼼸��");
        
        row=sheet.createRow(count);
        count++;
        
        // ��� ���� ����
        cell=row.createCell(0);
        cell.setCellValue("�Ͻ�");
        
        cell=row.createCell(1);
        cell.setCellValue("��� ��");
        
        cell=row.createCell(2);
        cell.setCellValue("�ܼ� ��");
        
        cell=row.createCell(3);
        cell.setCellValue("ī�� ���̵�");
        
        cell=row.createCell(4);
        cell.setCellValue("���");
        
        cell=row.createCell(5);
        cell.setCellValue("����� ��");
        
        cell=row.createCell(6);
        cell.setCellValue("�ٹ� ȸ��");
        
        cell=row.createCell(7);
        cell.setCellValue("�ٹ� �μ�");
        
        cell=row.createCell(8);
        cell.setCellValue("����");
        

        
        // ����Ʈ�� size ��ŭ row�� ����
        
        //3���̻���� 1���� ����� �Ʒ� ��ǥ�õǰ� �ؾ���
        Student vo;
        for(int rowIdx=0; rowIdx < student.length; rowIdx++) {

        	for(int j=0;j<student[rowIdx].date.size();j++) {
        		
            vo = student[rowIdx];
	            if(vo.penalty>=3) {

	            
	        	// �� ����
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
        cell.setCellValue("���� 3���̸� �����");
        
        row=sheet.createRow(count);
        count++;
        cell=row.createCell(4);
        cell.setCellValue("���");
        
        cell=row.createCell(5);
        cell.setCellValue("�̸�");
        
        cell=row.createCell(6);
        cell.setCellValue("����");
        
        cell=row.createCell(7);
        cell.setCellValue("����");

        
        for(int i=0;i<student.length;i++) {
        	for(int j=0;j<student[i].date.size();j++) {
        		if(student[i].penalty>0&&student[i].penalty<3) {
        			row=sheet.createRow(count);
        			
        			cell = row.createCell(4);
        			cell.setCellValue(student[i].number.get(j));
        			
                    cell = row.createCell(5);
        			cell.setCellValue(student[i].username.get(j));
        			
        			cell = row.createCell(6);
        			cell.setCellValue(student[i].penalty+" ��");
        			
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
        
        
        
        // �Էµ� ���� ���Ϸ� ����
        File file = new File("C:\\Users\\user\\Desktop\\���߰����ڷ�\\dom2\\student.xlsx");
        FileOutputStream fos = null;
        
        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            System.out.println("���� : student ���� ������ �ݰ� �ٽ� �������ּ���.");
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
