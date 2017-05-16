package Excel.testpoi;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class testpoi implements ActionListener{

	private JFrame frame;
	private JTable table;
	private JLabel lbMSG;
	private JButton btnExcel;
	private JButton btnSearch;
	private String[] saTit = new String[] {"번호","성명"};
	private int[] iaCwidth = new int[] {30, 70};
	private int[] iaAlm = new int[] {JLabel.RIGHT,JLabel.LEFT};
	private DefaultTableModel dtModel;
	String driver = "oracle.jdbc.driver.OracleDriver";
	String url = "jdbc:oracle:thin:@localhost:1521:XE";
	String user = "madang";
	String password = "madang";
	Connection con = null;
	PreparedStatement pstmt = null;
	ResultSet rs = null;
	ArrayList<String[]> backup = null;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					testpoi window = new testpoi();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public testpoi() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		btnExcel = new JButton("Excel 출력");
		btnExcel.setBounds(300, 218, 97, 23);
		frame.getContentPane().add(btnExcel);
		btnExcel.addActionListener(this);

		btnSearch = new JButton("조회");
		btnSearch.setBounds(303, 10, 97, 23);
		frame.getContentPane().add(btnSearch);
		btnSearch.addActionListener(this);

		lbMSG = new JLabel("");
		lbMSG.setBounds(58, 222, 163, 15);
		frame.getContentPane().add(lbMSG);

		JLabel lblNewLabel_4 = new JLabel("msg");
		lblNewLabel_4.setBounds(12, 222, 57, 15);
		frame.getContentPane().add(lblNewLabel_4);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(12, 38, 391, 170);
		frame.getContentPane().add(scrollPane);

		dtModel = new DefaultTableModel(saTit, 0);
		table = new JTable();
		scrollPane.setColumnHeaderView(table);
		table.setAutoCreateColumnsFromModel(false);
		table.setModel(dtModel);

		for(int i = 0 ; i < iaAlm.length ; i++)
		{
			DefaultTableCellRenderer renderer = new DefaultTableCellRenderer();
			renderer.setHorizontalAlignment(iaAlm[i]);
			TableColumn column = new TableColumn(i, iaCwidth[i],renderer, null);
			table.addColumn(column);
		}
		table.setFocusable(false);
		scrollPane.setViewportView(table);
	}

	public void actionPerformed(ActionEvent e) {
		if(e.getSource()==btnSearch){
			search();
		}
		else if(e.getSource()==btnExcel){
			makeExcelFile();
		}

	}

	public void dbconnect(){
		try{
			Class.forName(driver);
			con = DriverManager.getConnection(url,user,password);
		}catch(Exception e){
			e.printStackTrace();
		}
	}

	public void search(){
		String sql = "SELECT * FROM test_customer";
		try{
			lbMSG.setForeground(Color.BLUE);
			lbMSG.setText("조회중...");
			dbconnect();
			pstmt = con.prepareStatement(sql);
			rs = pstmt.executeQuery();
			if(backup!=null)
				backup.clear();
			else
				backup = new ArrayList<String[]>();
				
				String saData[] = null;
				int i = 0;
				DefaultTableModel model = (DefaultTableModel)table.getModel();
				model.setNumRows(0);
				while(rs.next()){
					i++;
					saData = new String[2];
					saData[0] = String.valueOf(rs.getInt(1));
					saData[1] = rs.getString(2);
					dtModel.addRow(saData);
					backup.add(saData);
				}
				if(i == 0){
					lbMSG.setForeground(Color.RED);
					lbMSG.setText("해당 코드가 없습니다.");
				}else{
					lbMSG.setForeground(Color.BLACK);
					lbMSG.setText("조회 완료. ");
				}

//				String[] s;
//				for(int j = 0 ; j < backup.size() ; j++) {
//					s = backup.get(j);
//					System.out.println(s[0] + " , " + s[1]);
//				}
		}catch(Exception e){
			lbMSG.setForeground(Color.RED);
			lbMSG.setText("Exception 발생");
			e.printStackTrace();
		}finally{
			try{if(rs != null) rs.close();
			}catch(SQLException e1){}
			try{if(pstmt != null) pstmt.close();
			}catch(SQLException e1){}
			try{if(con != null) con.close();
			}catch(SQLException e1){}
		}
	}
	public void makeExcelFile(){
		lbMSG.setForeground(Color.BLUE);
		lbMSG.setText("생성중...");
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet1 = wb.createSheet("첫장");
		sheet1.setColumnWidth(0, 10);
		XSSFRow row = sheet1.createRow(0);
		XSSFCell cell;
		String[] str = null;

		cell = row.createCell(1);
		cell.setCellValue("번호");
		cell = row.createCell(2);
		cell.setCellValue("이름");

		for(int i = 0 ; i < backup.size() ; i++){
			str = backup.get(i);
			row = sheet1.createRow(i+1);

			cell = row.createCell(1);
			cell.setCellValue(str[0]);
			cell = row.createCell(2);
			cell.setCellValue(str[1]);
		}
		File file = new File("D:\\testpoi.xlsx");
		FileOutputStream fileout = null;
		try{
			fileout = new FileOutputStream(file);
			wb.write(fileout);
			lbMSG.setForeground(Color.BLACK);
			lbMSG.setText("엑셀 파일 생성 완료.");
		}catch(FileNotFoundException e) {
			lbMSG.setForeground(Color.RED);
			lbMSG.setText("FileNotFoundException 발생");
			e.printStackTrace();
		}catch(IOException e){
			lbMSG.setForeground(Color.RED);
			lbMSG.setText("IOException 발생");
			e.printStackTrace();
		}catch(Exception e){
			lbMSG.setForeground(Color.RED);
			lbMSG.setText("Exception 발생");
			e.printStackTrace();
		}finally{
			try{if(wb!=null) wb.close();
			}catch(IOException e1){}
			try{if(fileout!=null) fileout.close();
			}catch(IOException e1){}
		}
	}
}

