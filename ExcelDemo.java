package package1;

import java.awt.BorderLayout;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.MouseMotionListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.channels.NetworkChannel;
import java.util.StringTokenizer;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumnModel;

public class ExcelDemo extends JFrame{
	//클래스의 멤버들을 설정
	private JScrollPane scrollPane;
	private JTable table, headerTable;
	private JMenuBar menuBar;
	private JMenu fileMenu, formulasMenu, functionMenu;
	private JMenuItem newItem, open, save, exit, sum, average, count, max, min;
	private String title;
	private int cardinality, degree;
	
	//생성자 정의
	public ExcelDemo() {
		String headerColumn[] = new String[26];
		char A = 'A';
		for(int i = 0; i< 26; i++) {
			headerColumn[i] = Character.toString(A);
			A++;
		}//헤더컬럼을 만들어 ABC순으로 안에 집어넣음. 
		String contents[][] = new String[100][26];
		for(int i = 0; i < 100; i++) {
			for(int j = 0; j < 26; j++) {
				contents[i][j] = " ";
			}
		}
		table = new JTable(contents, headerColumn);//100x26 배열과 헤더컬럼을 테이블에 넣음.
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);//테이블 셀 사이즈를 이쁘게 만듦.
		
		//테이블의 셀을 마우스로 선택했을때의 행동을 정의
		table.addMouseListener(new MouseListener() {
			
			@Override
			public void mouseReleased(MouseEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void mousePressed(MouseEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void mouseExited(MouseEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void mouseEntered(MouseEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			//셀을 마우스 클릭을 했을때 cardinality , degree 에 각각 해당셀의 row값과 column값을 넣음.
			@Override
			public void mouseClicked(MouseEvent e) {
				// TODO Auto-generated method stub
				if(e.getButton() == 1) {
					cardinality = table.getSelectedRow();
					degree = table.getSelectedColumn();
					
				}
			}
		});
		
		//로우헤더의 값을 변경 불가로 만들기 위한 사전 작업
		DefaultTableModel model = new DefaultTableModel(100,1) {
			@Override
			public boolean isCellEditable(int row, int column) {
				return false;
			}
		};
		headerTable = new JTable(model);
	
		for(int i = 0; i< 100; i++) {
			headerTable.setValueAt(i, i, 0);//로우헤더의 각 셀에 0~99를 집어넣음.
		}
		

		headerTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);//로우헤더의 셀 크기를 이쁘게 만듦.
		headerTable.getColumnModel().getColumn(0).setPreferredWidth(50);
	    headerTable.setPreferredScrollableViewportSize(new Dimension(50,0));//테이블과 헤더테이블 사이의 공백을 없애줌.
	    headerTable.setBackground(table.getTableHeader().getBackground());//헤더테이블의 색상을 테이블의 헤더의 색상과 같게 함.
	    headerTable.setFocusable(false);//헤더테이블 셀 더블클릭 안되도록 설정
	    headerTable.setSelectionModel(table.getSelectionModel());//셀 선태갛면 그 row header색칠되도록
	    
	    //셀 선택시 같은 행의 로우헤더 숫자를 진하게 표시
	    DefaultTableCellRenderer render = new DefaultTableCellRenderer() { 
			public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
				JComponent c = (JComponent) super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			
				if(isSelected)
					c.setFont(getFont().deriveFont(Font.BOLD));
				
				return c;
			}
		};
		render.setHorizontalAlignment(SwingConstants.CENTER);
	    headerTable.setDefaultRenderer(headerTable.getColumnClass(0), render);
	    scrollPane = new JScrollPane(table);//스크롤 구현
	    
		
		JFrame frame = new JFrame("새 Microsoft Excel 워크시트.xlsx - Excel");//프레임을 만듦
		frame.setSize(610, 588);//프레임 크기 설정
		
		//프레임을 화면 가운데 정렬
		Dimension frameSize = frame.getSize();//프레임크기
	    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();//모니터크기
	    frame.setLocation((screenSize.width - frameSize.width) / 2, (screenSize.height - frameSize.height) / 2);
	    
	    frame.setLayout(new BorderLayout());//보더레이아웃으로 설정.
		
	    //테이블과 헤더테이블을 프레임에 집어넣음. 헤더테이블도 스크롤 호환 가능하게 만듦.
	    frame.add(scrollPane);
		frame.add(headerTable,BorderLayout.WEST);
		scrollPane.setRowHeaderView(headerTable);
		
		//메뉴를 프레임에 구현시킴.
		createMenu(frame);
		frame.setJMenuBar(menuBar);
		
		//프레임을 보이게 만들고, x를 누르면 메모리 소멸되게 함.
		frame.setVisible(true);
		frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
		
	}
	
	//생성자 안에 메뉴 관련 객체들을 만들기에는 너무 코드가 길어서 따로 함수를 정의하여 메뉴관련 객체들을 만듦.
	public void createMenu(JFrame frame) {
		
		//메뉴바에 file, formulas 메뉴 버튼 나오게 함.
		menuBar = new JMenuBar();
		fileMenu = new JMenu("File");
		menuBar.add(fileMenu);
		formulasMenu = new JMenu("Formulas");
		menuBar.add(formulasMenu);
		
		//new 메뉴아이템 정의, new 클릭 시 창이 닫히고, 새 창이 열리는 거 리스너를 통해 구현.
		newItem = new JMenuItem("New");
		newItem.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				new ExcelDemo();
			}
		});
		fileMenu.add(newItem);
		
		//open 메뉴 아이템 정의, open 클릭 시 발생되는 것들 리스너로 구현.
		open = new JMenuItem("Open");
		open.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser("Open");
				FileNameExtensionFilter filter = new FileNameExtensionFilter("txt & csv", "txt", "csv");//txt, csv 형식의 파일만을 필터링해줌.
				chooser.setFileFilter(filter);
				
				int ret = chooser.showOpenDialog(null);//오픈 다이얼 로그를 출력
				
				//파일을 선택하지 않았을 경우 걸러줌.
				if(ret!=JFileChooser.APPROVE_OPTION){
					return;
		        }
			
				String filePath = chooser.getSelectedFile().getPath();
				frame.setTitle(filePath);//프레임 타이틀을 파일 경로로 변경.
				
				//프레임 셀 모두 초기화
				for(int i=0;i<26;i++) {
	                  for (int j=0;j<100;j++) {
	                     table.setValueAt(" ", j, i);
	                  }
	               }
	               
	               //파일에 있는 문자들을 읽어 프레임의 각 셀에 적음.
	               int row=0;
	               int col=0;
	               try {
	                  FileReader FR=new FileReader(chooser.getSelectedFile());
	                  BufferedReader BR=new BufferedReader(FR);
	                  
	                  String line = null;
	                  String cell = null;
	                  
	                  //파일의 마지막 줄까지 실행.
	                  while((line=BR.readLine()) != null) {
	                      
	                	  //파일의 한 줄을 stringTokenizer가 받아들이기 쉽게 가공하여 temp에 넣음. 예) 1,2,3,,4,,,5 -> 1,2,3, ,4, , ,5
	                	  String temp="";
	                     String[] split=line.split(",");
	                     for(int i=0;i<split.length;i++) {
	                        if(split[i].equals("")) {
	                           temp=temp+" "+",";
	                        }
	                        else {
	                           temp=temp+split[i]+",";
	                        }
	                     }
	                     
	                     //가공된 temp를 테이블 셀에 "," 단위로 끊어서 각각 집어넣음.
	                     StringTokenizer ST =new StringTokenizer(temp,",");
	                     col = 0;
	                     for(;ST.hasMoreTokens();col++) {
	                        cell =ST.nextToken();//다음 token
	                        table.setValueAt(cell, row, col);//값을 table에 추가
	                        
	                     }
	                     row++;
	                 }
	                 //파일 닫기 
	                 FR.close();
	                 BR.close();
	               }
	               catch(FileNotFoundException e) {
	            	   return;
	               }
	               catch(IOException E) {
	            	  return;
	               }
   
				
				
			}
		});
		fileMenu.add(open);
		
		//세이브 메뉴 아이템 정의. 세이브 메뉴 클릭시 발생되는 것들 리스너로 구현.
		save = new JMenuItem("Save");
		fileMenu.add(save);
		save.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				int ret = chooser.showSaveDialog(null);//세이브 다이얼로그 화면 출력.
				
				//파일 선택하지 않았을 경우 걸러줌.
				if(ret!=JFileChooser.APPROVE_OPTION){
					return;
		        }

				//셀의 내용들을 파일에 집어넣음.
				try {
					FileWriter fw = new FileWriter(chooser.getSelectedFile().getPath() + ".txt");
					BufferedWriter bw = new BufferedWriter(fw);
					
					//셀의 경계를 ","로 표현하여 파일에 집어넣음.
					for(int i=0;i<100;i++) {
		                  for (int j=0;j<26;j++) {
		                	
		                	 if(table.getValueAt(i, j) !=null) {
		                		 bw.write(table.getValueAt(i,j).toString()+ ",");
		                	 }
		                	 else {
		                		 bw.write(",");
		                	 }
		                  }
		                  bw.flush();
		                  bw.newLine();
					}
					
					//파일 닫기
					bw.close();
					fw.close();
				}catch(Exception e) {
					return;
				}
			}
		});
		
		//exit 메뉴 아이템 정의. exit 메뉴 아이템 클릭 시 종료되는 것 리스너로 구현.
		exit = new JMenuItem("Exit");
		fileMenu.add(exit);
		exit.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);//종료되게 만듦.
			}
		});
		
		//function의 각 메뉴 아이템들을 구현. 
		functionMenu = new JMenu("Function");
		formulasMenu.add(functionMenu);
		
		//sum 메뉴아이템 정의. 리스너로 구현.
		sum = new JMenuItem("SUM");
		sum.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int column1, column2, row1, row2;
				double sum = 0;
				String row1Str = "";
				String row2Str = "";
				
				//쇼다이얼로그를 화면에 띄우고 문자를 입력받음.
				String str = JOptionPane.showInputDialog(null, "Function Arguments", "SUM", JOptionPane.PLAIN_MESSAGE);
				
				//만약 취소를 누르거나 아무것도 입력하지 않았을때에 예외가 발생하지 않도록 if문으로 분기해준다.
				if(str == null || str.equals("")) {
					return;
				}
				StringTokenizer stn = new StringTokenizer(str, ":");
				
				//":"로 나뉘어진 각 문자열의 첫번째 요소를 숫자로 변환시킨 후, column에 넣음. 예) A0:B2 -> A -> 0를 column1에 B -> 1를 column2에 넣음.
				String first = stn.nextToken();
				String second = stn.nextToken();
				column1 = first.charAt(0) - 65;
				column2 = second.charAt(0) - 65;
				
				//":"로 나뉘어진 각 문자열의 두번째 요소부터 마지막까지의 문자들을 String형식의 문자열에 집어넣음. 예) A12:D20 --> "12"를 row1Str에, "20"를 row2Str에 넣음.
				for(int i = 1; i < first.length();i++) {
					row1Str += Character.toString(first.charAt(i));
				}
				for(int i = 1; i < second.length();i++) {
					row2Str += Character.toString(second.charAt(i));
				}
				//String형식을 int형식으로 바꿈.
				row1 = Integer.parseInt(row1Str);
				row2 = Integer.parseInt(row2Str);
				
				//위에서 구한 row와 column들로 셀의 구간을 정한 뒤, sum을 구함.
				for(int i =row1; i <= row2; i++) {
					for(int j = column1; j <= column2; j++) {
						if(table.getValueAt(i, j).toString().equals(" ") == false) {
							sum += Double.parseDouble(table.getValueAt(i,j).toString());
						}
					}
				}
				table.setValueAt(sum+"", cardinality, degree);//sum을 마우스로 클릭한 곳에 출력시킴.
			}
		});
		functionMenu.add(sum);
		
		//average 메뉴 아이템 정의. 리스너로 구현. sum과 매우 유사하기 때문에 동일한 역할을 하는 코드들은 주석 처리하지 않았음. 다만 출력시키는 값을 count로 나눠주고 출력시키면 됨.
		average = new JMenuItem("AVERAGE");
		functionMenu.add(average);
		average.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int column1, column2, row1, row2;
				double average = 0;
				double sum = 0;
				int count = 0;//sum을 count로 나누어 average 구함.
				
				String row1Str = "";
				String row2Str = "";
				
				String str = JOptionPane.showInputDialog(null, "Function Arguments", "AVERAGE", JOptionPane.PLAIN_MESSAGE);
				
				if(str == null || str.equals("")) {
					return;
				}
				
				StringTokenizer stn = new StringTokenizer(str, ":");
				
				String first = stn.nextToken();
				String second = stn.nextToken();
				column1 = first.charAt(0) - 65;
				column2 = second.charAt(0) - 65;
				
				for(int i = 1; i < first.length();i++) {
					row1Str += Character.toString(first.charAt(i));
				}
				for(int i = 1; i < second.length();i++) {
					row2Str += Character.toString(second.charAt(i));
				}
				row1 = Integer.parseInt(row1Str);
				row2 = Integer.parseInt(row2Str);
				
				for(int i =row1; i <= row2; i++) {
					for(int j = column1; j <= column2; j++) {
						if(table.getValueAt(i, j).toString().equals(" ") == false) {
							sum += Double.parseDouble(table.getValueAt(i,j).toString());
							count++;//비어있지 않은 셀일 때마다 1씩 올려준다.
						}
					}
				}
				
				//count = 0 일때 나누어주면 안되므로 if문으로 걸러주었다.
				if(count != 0) {
					average = sum/count;
				}
				table.setValueAt(average+"", cardinality, degree);
			}
			
		});
		
		//count 메뉴 아이템 정의. 위의 sum, average와 동일한 기능을 하는 코드들의 주석은 처리하지 않았다. 다만 출력시키는 값은 해당 셀 영역의 비어있지 않은 셀의 갯수이다.
		count = new JMenuItem("COUNT");
		functionMenu.add(count);
		count.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int column1, column2, row1, row2;
				int count = 0;
				
				String row1Str = "";
				String row2Str = "";
				
				String str = JOptionPane.showInputDialog(null, "Function Arguments", "COUNT", JOptionPane.PLAIN_MESSAGE);
				
				if(str == null || str.equals("")) {
					return;
				}
				
				StringTokenizer stn = new StringTokenizer(str, ":");
				
				String first = stn.nextToken();
				String second = stn.nextToken();
				column1 = first.charAt(0) - 65;
				column2 = second.charAt(0) - 65;
				
				for(int i = 1; i < first.length();i++) {
					row1Str += Character.toString(first.charAt(i));
				}
				for(int i = 1; i < second.length();i++) {
					row2Str += Character.toString(second.charAt(i));
				}
				row1 = Integer.parseInt(row1Str);
				row2 = Integer.parseInt(row2Str);
				
				for(int i =row1; i <= row2; i++) {
					for(int j = column1; j <= column2; j++) {
						if(table.getValueAt(i, j).toString().equals(" ") == false) {
							count++;
						}
					}
				}
				table.setValueAt(count+"", cardinality, degree);//count값을 마우스로 클릭한 셀에 출력.
			}
			
		});
		
		//max 메뉴 아이템 정의. 위의 sum, average, count와 동일한 기능을 하는 코드들의 주석은 처리하지 않음. 다만 출력시키는 값은 해당 셀 영역 중 가장 큰 값이다.
		max= new JMenuItem("MAX");
		functionMenu.add(max);
		max.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int column1, column2, row1, row2;
				String row1Str = "";
				String row2Str = "";
				
				String str = JOptionPane.showInputDialog(null, "Function Arguments", "MAX", JOptionPane.PLAIN_MESSAGE);
				
				if(str == null || str.equals("")) {
					return;
				}
				
				StringTokenizer stn = new StringTokenizer(str, ":");
				
				String first = stn.nextToken();
				String second = stn.nextToken();
				column1 = first.charAt(0) - 65;
				column2 = second.charAt(0) - 65;
				
				for(int i = 1; i < first.length();i++) {
					row1Str += Character.toString(first.charAt(i));
				}
				for(int i = 1; i < second.length();i++) {
					row2Str += Character.toString(second.charAt(i));
				}
				row1 = Integer.parseInt(row1Str);
				row2 = Integer.parseInt(row2Str);
				
				double max = Double.parseDouble(table.getValueAt(row1, column1).toString());
				double num;
				for(int i =row1; i <= row2; i++) {
					for(int j = column1; j <= column2; j++) {
						//더 큰 값을 마주했을 때에 max의 값을 해당 값으로 변경한다.
						if(table.getValueAt(i, j).toString().equals(" ") == false) {
							num = Double.parseDouble(table.getValueAt(i, j).toString());
							if(max < num) {
								max = num;
							}
						}
					}
				}
				table.setValueAt(max+"", cardinality, degree);//마우스로 클릭한 해당 셀에 max값을 출력한다.
			}
			
		});
		
		//min 메뉴 아이템 정의. 위의 코드들과 유사한 기능을 하는 코드들의 주석은 처리하지 않았다. 다만 출력시키는 값은 해당 셀 영역 중 가장 값이 작은 값이다.
		min = new JMenuItem("MIN");
		functionMenu.add(min);
		min.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int column1, column2, row1, row2;
				String row1Str = "";
				String row2Str = "";
				
				String str = JOptionPane.showInputDialog(null, "Function Arguments", "MIN", JOptionPane.PLAIN_MESSAGE);
				
				if(str == null || str.equals("")) {
					return;
				}
				
				StringTokenizer stn = new StringTokenizer(str, ":");
				
				String first = stn.nextToken();
				String second = stn.nextToken();
				column1 = first.charAt(0) - 65;
				column2 = second.charAt(0) - 65;
				
				for(int i = 1; i < first.length();i++) {
					row1Str += Character.toString(first.charAt(i));
				}
				for(int i = 1; i < second.length();i++) {
					row2Str += Character.toString(second.charAt(i));
				}
				row1 = Integer.parseInt(row1Str);
				row2 = Integer.parseInt(row2Str);
				
				double min = Double.parseDouble(table.getValueAt(row1, column1).toString());
				double num;
				for(int i =row1; i <= row2; i++) {
					for(int j = column1; j <= column2; j++) {
						//min보다 작은 값을 마주했을 때에 해당 값으로 min의 값을 변경.
						if(table.getValueAt(i, j).toString().equals(" ") == false) {
							num = Double.parseDouble(table.getValueAt(i, j).toString());
							if(min > num) {
								min = num;
							}
						}
					}
				}
				table.setValueAt(min+"", cardinality, degree);//마우스로 클릭한 해당 셀에 min값 출력.
			}	
		});
		
	}
	
	//메인문
	public static void main(String[] args) {
		ExcelDemo excel = new ExcelDemo();
			
	}
}

/*






*/
