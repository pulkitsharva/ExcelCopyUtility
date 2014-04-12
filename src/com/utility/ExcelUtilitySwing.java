package com.utility;


import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.Font;
import javax.swing.SwingConstants;
import java.awt.Color;
import java.awt.SystemColor;

public class ExcelUtilitySwing extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JTextField sourceField;
	private JTextField destinationField;
	private JTextField sourceField1;
	private JTextField sourceField2;
	private JTextField sourceField3;
	
	public static String currentDir;
	public static String destinationDir;
	public static String resultFileName;
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args)
	{
		EventQueue.invokeLater(new Runnable() 
		{
			public void run()
			{
				try
				{
					ExcelUtilitySwing frame = new ExcelUtilitySwing();
					frame.setVisible(true);
				}
				catch (Exception e1) 
				{
					System.out.println(e1.getLocalizedMessage());
					e1.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public ExcelUtilitySwing() {
		setTitle("Excel Copy Utility\r\n");
		setForeground(SystemColor.activeCaption);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(500, 100, 401, 320);
		setLocationRelativeTo(null);
		getContentPane().setLayout(null);
		
		sourceField = new JTextField();
		sourceField.setBounds(134, 71, 195, 20);
		getContentPane().add(sourceField);
		sourceField.setColumns(10);
		
		sourceField1 = new JTextField();
		sourceField1.setBounds(134, 101, 195, 20);
		getContentPane().add(sourceField1);
		sourceField1.setColumns(10);
		//sourceField.setEditable(false);
		
		sourceField2 = new JTextField();
		sourceField2.setBounds(180, 131, 150, 20);
		getContentPane().add(sourceField2);
		sourceField2.setColumns(10);
		
		sourceField3 = new JTextField();
		sourceField3.setBounds(134, 161, 195, 20);
		getContentPane().add(sourceField3);
		sourceField3.setColumns(10);
		
		
		destinationField = new JTextField();
		destinationField.setBounds(134, 191, 195, 20);
		getContentPane().add(destinationField);
		//destinationField.setEditable(false);
		destinationField.setColumns(10);
		
		JButton btnChooseFile = new JButton("Round 1");
		btnChooseFile.setHorizontalAlignment(SwingConstants.LEFT);
		btnChooseFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 JFileChooser c = new JFileChooser();
			      // Demonstrate "Open" dialog:
					  int rVal = c.showOpenDialog(ExcelUtilitySwing.this);
					  if (rVal == JFileChooser.APPROVE_OPTION) 
					  {
						  
						  sourceField.setText(c.getCurrentDirectory().toString()+"\\"+c.getSelectedFile().getName());
					  }
					  if (rVal == JFileChooser.CANCEL_OPTION) {
					       sourceField.setText("");
					      }
				  }
			});
		btnChooseFile.setBounds(23, 70, 88, 23);
		getContentPane().add(btnChooseFile);
		
		
		JButton btnChooseFile1 = new JButton("Round 2");
		btnChooseFile1.setHorizontalAlignment(SwingConstants.LEFT);
		btnChooseFile1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 JFileChooser c = new JFileChooser();
			      // Demonstrate "Open" dialog:
					  int rVal = c.showOpenDialog(ExcelUtilitySwing.this);
					  if (rVal == JFileChooser.APPROVE_OPTION) 
					  {
						  
						  sourceField1.setText(c.getCurrentDirectory().toString()+"\\"+c.getSelectedFile().getName());
					  }
					  if (rVal == JFileChooser.CANCEL_OPTION) {
					       sourceField1.setText("");
					      }
					  
				  }
			
		});
		
		btnChooseFile1.setBounds(23, 100, 88, 23);
		getContentPane().add(btnChooseFile1);
		
		
		JButton btnChooseFile2 = new JButton("Lookup Transaction");
		btnChooseFile2.setHorizontalAlignment(SwingConstants.LEFT);
		btnChooseFile2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 JFileChooser c = new JFileChooser();
			      // Demonstrate "Open" dialog:
					  int rVal = c.showOpenDialog(ExcelUtilitySwing.this);
					  if (rVal == JFileChooser.APPROVE_OPTION) 
					  {
						  sourceField2.setText(c.getCurrentDirectory().toString()+"\\"+c.getSelectedFile().getName());
					  }
					  if (rVal == JFileChooser.CANCEL_OPTION) {
					       sourceField2.setText("");
					      }
					  }
			});
		btnChooseFile2.setBounds(23, 130, 148, 23);
		getContentPane().add(btnChooseFile2);
		
		
		JButton btnChooseFile3 = new JButton("Template");
		btnChooseFile3.setHorizontalAlignment(SwingConstants.LEFT);
		btnChooseFile3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 JFileChooser c = new JFileChooser();
			      // Demonstrate "Open" dialog:
					  int rVal = c.showOpenDialog(ExcelUtilitySwing.this);
					  if (rVal == JFileChooser.APPROVE_OPTION) 
					  {
						  sourceField3.setText(c.getCurrentDirectory().toString()+"\\"+c.getSelectedFile().getName());
					  }
					  if (rVal == JFileChooser.CANCEL_OPTION) {
					       sourceField3.setText("");
					      }
					  }
			});
		btnChooseFile3.setBounds(23, 160,88, 23);
		getContentPane().add(btnChooseFile3);
		
		
		JButton btnDestination = new JButton("Result");
		btnDestination.setHorizontalAlignment(SwingConstants.CENTER);
		btnDestination.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) 
			{
				JFileChooser c = new JFileChooser();
			      // Demonstrate "Save" dialog:
			      int rVal = c.showSaveDialog(ExcelUtilitySwing.this);
			      if (rVal == JFileChooser.APPROVE_OPTION) {
			    	  destinationDir=c.getCurrentDirectory().toString();
			    	  resultFileName=c.getCurrentDirectory().toString()+"\\"+c.getSelectedFile().getName()+".xls";
			        destinationField.setText(resultFileName);
			        }
			      if (rVal == JFileChooser.CANCEL_OPTION) {
			        destinationField.setText("");
			      }
			}
		});
		btnDestination.setBounds(23, 190, 88, 23);
		getContentPane().add(btnDestination);
		
		
		
		JButton btnConvert = new JButton("Convert");
		btnConvert.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) 
			{
				try
				{
					if(sourceField.getText().equals(""))
					{
						AlertBox.alert("Please provide Round 1 Raw Data file.");
					}
					else if(sourceField1.getText().equals(""))
					{
						AlertBox.alert("Please provide Round 2 Raw Data file.");
					}
					else if(sourceField2.getText().equals(""))
					{
						AlertBox.alert("Please provide Lookup Transaction file.");
					}
					else if(sourceField3.getText().equals(""))
					{
						AlertBox.alert("Please provide Blank Template file.");
					}
					else if(destinationField.getText().equals(""))
					{
						AlertBox.alert("Please provide Destination for Result file.");
					}
					else
					{
						ExcelFormat excelFormat=new ExcelFormat();
						excelFormat.copyRawExcel(sourceField.getText(), destinationField.getText(),true);
						excelFormat.copyRawExcel(sourceField1.getText(), destinationField.getText(),false);
						excelFormat.excelCopy(sourceField2.getText(),sourceField3.getText(),resultFileName);
						
						//excelFormat.excelCopy();
						AlertBox.alert("Excel copying done.");
					}
				}
				catch(Exception e2)
				{
					System.out.println(e2.getMessage());
					e2.printStackTrace();
				}
			}
		});
		btnConvert.setBounds(148, 240, 107, 31);
		getContentPane().add(btnConvert);
		
		JLabel lblPleaseSpecifyxls = new JLabel("Please specify .xls as destination file extension");
		lblPleaseSpecifyxls.setHorizontalAlignment(SwingConstants.LEFT);
		lblPleaseSpecifyxls.setForeground(Color.RED);
		lblPleaseSpecifyxls.setBounds(109, 217, 241, 14);
		getContentPane().add(lblPleaseSpecifyxls);
		
		JLabel lblExcelCopyUtility = new JLabel("Excel Copy Utility");
		lblExcelCopyUtility.setHorizontalAlignment(SwingConstants.CENTER);
		lblExcelCopyUtility.setFont(new Font("Times New Roman", Font.BOLD, 16));
		lblExcelCopyUtility.setBounds(10, 11, 365, 31);
		getContentPane().add(lblExcelCopyUtility);
	}
}
