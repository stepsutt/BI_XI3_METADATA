package com.ibm.bi.tools;

import java.awt.GridLayout;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.JLabel;
import javax.swing.JPasswordField;
import javax.swing.JOptionPane;

public class GetBILogon {
	
	public String strUser;
	public String strPassword;
	public String strCMS;
	
	public GetBILogon () {
		strUser = "";
		strPassword ="";
		strCMS = "";
	}
	
	public int getUserName() {
		
		boolean bOK = false;
		int result = 0;
		
	    JTextField User = new JTextField("");
	    JTextField Password = new JPasswordField("");
	    JTextField CMS = new JTextField("");
	    JPanel panel = new JPanel(new GridLayout(0, 1));
	    panel.add(new JLabel("User:"));
	    panel.add(User);
	    panel.add(new JLabel("Password:"));
	    panel.add(Password);
	    panel.add(new JLabel("CMS:"));
	    panel.add(CMS);
	    while (!bOK) {
	    	try {
	    		result = JOptionPane.showConfirmDialog(null, panel, "IBM SAP BI Tool",
	    			JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
	    		if (result == JOptionPane.CANCEL_OPTION) {
	    			bOK = true;
	    		} else if (Password.getText().equals("")) {
	    			JOptionPane.showMessageDialog(null, "Must enter a password");	    			
	    		} else if (User.getText().equals("")) {
	    			JOptionPane.showMessageDialog(null, "Must enter a user");	    			
	    		} else if (CMS.getText().equals("")) {
	    			JOptionPane.showMessageDialog(null, "Must enter a CMS");	    			
	    		} else {
	    			bOK = true;
	    			strUser = User.getText();
	    			strPassword = Password.getText();
	    			strCMS = CMS.getText();
	    		}
	    	} catch (Exception e) {
	    		System.out.println(e.toString());
	    	}
	    }
	    
	    
	    return result;
	}

}

