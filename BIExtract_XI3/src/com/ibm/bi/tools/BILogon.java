package com.ibm.bi.tools;

import java.awt.EventQueue;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.JCheckBox;
import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import org.eclipse.wb.swing.FocusTraversalOnArray;
import java.awt.Component;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class BILogon extends JFrame {

	private JPanel contentPane;
	private JTextField txtUser;
	private JPasswordField passPassword;
	private JTextField txtCMS;
	private JCheckBox chckbxDebugMode;
	private final ButtonGroup buttonGroup = new ButtonGroup();
	private boolean isSubmit = false;

	JButton btnLogon;
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					BILogon frame = new BILogon();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public BILogon() {
		setTitle("IBM SAP BO Extract");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 365, 263);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JLabel lblUser = new JLabel("User:");
		lblUser.setBounds(91, 14, 46, 14);
		contentPane.add(lblUser);
		
		txtUser = new JTextField();
		lblUser.setLabelFor(txtUser);
		txtUser.setBounds(157, 11, 86, 20);
		contentPane.add(txtUser);
		txtUser.setColumns(10);
		
		JLabel lblPassword = new JLabel("Password:");
		lblPassword.setBounds(91, 52, 57, 14);
		contentPane.add(lblPassword);
		
		passPassword = new JPasswordField();
		lblPassword.setLabelFor(passPassword);
		passPassword.setBounds(157, 49, 86, 20);
		contentPane.add(passPassword);
		
		JLabel lblCms = new JLabel("CMS:");
		lblCms.setBounds(91, 90, 46, 14);
		contentPane.add(lblCms);
		
		txtCMS = new JTextField();
		lblCms.setLabelFor(txtCMS);
		txtCMS.setBounds(157, 87, 86, 20);
		contentPane.add(txtCMS);
		txtCMS.setColumns(10);
		
		chckbxDebugMode = new JCheckBox("Debug Mode?");
		chckbxDebugMode.setBounds(62, 126, 97, 23);
		contentPane.add(chckbxDebugMode);
		
		JRadioButton rdbtnDosWindow = new JRadioButton("DOS window");
		rdbtnDosWindow.setSelected(true);
		buttonGroup.add(rdbtnDosWindow);
		rdbtnDosWindow.setBounds(187, 126, 109, 23);
		contentPane.add(rdbtnDosWindow);
		
		JRadioButton rdbtnTextFile = new JRadioButton("Text file");
		buttonGroup.add(rdbtnTextFile);
		rdbtnTextFile.setBounds(187, 155, 109, 23);
		contentPane.add(rdbtnTextFile);
		
		btnLogon = new JButton("Logon");
		btnLogon.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.out.println(arg0.toString());
				isSubmit = true;
				dispose();
			}
		});
		btnLogon.setBounds(139, 191, 89, 23);
		contentPane.add(btnLogon);
		contentPane.setFocusTraversalPolicy(new FocusTraversalOnArray(new Component[]{lblPassword, txtUser, passPassword, lblUser, lblCms, txtCMS, chckbxDebugMode, rdbtnDosWindow, rdbtnTextFile, btnLogon}));
	}
	
	public boolean isOK() {
		return isSubmit;
	}
	
	public String getUsername() {
	        return txtUser.getText().trim();
	}
	 
	public String getPassword() {
	        return new String(passPassword.getPassword());
	}
	
	public String getCMS() {
			return txtCMS.getText().trim(); 
	}
	
	public boolean getDebug() {
			return chckbxDebugMode.isSelected();
	}
	
	public String getDebugOut() {
			return buttonGroup.getSelection().toString();
	}
	
}
