package com.ibm.bi.tools;

import java.text.SimpleDateFormat;
import java.util.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import com.crystaldecisions.sdk.framework.*;
import com.crystaldecisions.sdk.properties.*;
import com.crystaldecisions.sdk.occa.infostore.*;
import com.crystaldecisions.sdk.plugin.desktop.user.*;
//import com.ibm.sapbi.logon.*;
import com.ibm.util.excel.*;
import com.businessobjects.sdk.plugin.desktop.webi.*;
import com.businessobjects.sdk.plugin.desktop.universe.*;
import com.crystaldecisions.sdk.plugin.desktop.folder.*;

public class BIExtractXI3 {
	
	private static WriteToExcel wtExcel;
	//private static String curFolder = "";
	private static Calendar dtLocal = new GregorianCalendar();
	private static SimpleDateFormat dateFormatter = new SimpleDateFormat("dd-MMM-yyyy HH:mm:ss");
	private static int iLimit = 100000;
	private static JFrame frame;
	private static String[] arrStatus = {"Running","Complete","2","Failure","4","5","6","7","Paused","Pending"};
	private static IInfoStore iStore = null;
	private static long heapFreeSize = 0;
	private static GetBILogon mp = null;
	
	private static void currentTime() {

	    	Calendar cal = Calendar.getInstance();
	    	cal.getTime();
	    	SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
	    	System.out.println("+++++++++++++++ " +  sdf.format(cal.getTime()) );

	}
	
	private static void msgbox(String sTitle, String s, int iType){
		   JOptionPane.showMessageDialog(null, s, sTitle, iType);
		}
	
	private static void getHeap() {
		heapFreeSize = Runtime.getRuntime().freeMemory();
		System.out.println("**********   HEAP SIZE is " + heapFreeSize);
	}
	
	private static void listRootFolder(String iParID, String sSheet, String sTab) {
		
		String sFolds[] = new String[6];
		String query = "Select top " + iLimit +  " * from " + sTab + " where si_parentid = " + iParID;
		
		try {
			wtExcel.createSheet(sSheet);
			System.out.println("Getting " + sSheet + " with " + query);
			sFolds[0] = "ID";
			sFolds[1] = "Name";
			sFolds[2] = "Description";
			sFolds[3] = "Ownder_ID";
			sFolds[4] = "Owner";
			sFolds[5] = "BO_System";
			wtExcel.writeHeader(sSheet, sFolds);
			IInfoObjects subFolders = iStore.query(query);
			for (int j=0;j < subFolders.size(); j++) {
				IInfoObject iObj = (IInfoObject)subFolders.get(j);
				sFolds[0] = iObj.properties().getProperty("SI_ID").toString();
				sFolds[1] = iObj.properties().getProperty("SI_NAME").toString();
				sFolds[2] = iObj.properties().getProperty("SI_DESCRIPTION").toString();
				sFolds[3] = iObj.properties().getProperty("SI_OWNERID").toString();
				sFolds[4] = iObj.properties().getProperty("SI_OWNER").toString();
				sFolds[5] = mp.strCMS;
				wtExcel.writeSheet(sSheet, sFolds);
			}
		} catch (Exception e) {
			System.out.println(e.toString());
		}
	}
	
	/*
	private static void getSubFolders(String iParID, String sPath) {
		
		String sFolds[] = new String[2];
		IInfoObjects subFolders = null;
		String query = "Select top " + iLimit +  " * from ci_infoobjects where si_kind ='Folder' and si_parentid = " + iParID;
		
		try {
			subFolders = iStore.query(query);
			System.out.println("Getting " + subFolders.size() + " subfolders");
			for (int j=0;j < subFolders.size(); j++) {
				System.out.println("Subfolder " + (j + 1) + " of " + subFolders.size());
				IInfoObject iObj = (IInfoObject)subFolders.get(j);
				String sFullName = sPath + "\\" + iObj.properties().getProperty("SI_NAME");
				System.out.println("Processing Folder " + sFullName);
				sFolds[0] = sFullName;
				sFolds[1] = iObj.properties().getProperty("SI_ID").toString();
				wtExcel.writeSheet("Hierarchy", sFolds);
				getReports(iObj.properties().getProperty("SI_ID").toString(), sFullName);
				getSubFolders(iObj.properties().getProperty("SI_ID").toString(), sFullName);
			}
		} catch (Exception e) {
			System.out.println(e.toString());
		}
		System.out.println("Got all subfolders");
	}
	*/
	
	/*
	private static void getInstances(String iParID, String sPath) {
	
		String[] rowData = new String[11];	
		IInfoObject iObj = null;
		
		String query = "Select top " + iLimit +  " * from ci_infoobjects where si_instance = 1 and si_parentid = " + iParID;
		try {
			IInfoObjects subFolders = iStore.query(query);
			for (int j=0;j < subFolders.size(); j++) {
				iObj = (IInfoObject)subFolders.get(j);
				rowData[0] = sPath;
				rowData[1] = iObj.properties().getProperty("SI_NAME").toString().replace("%26", "&").replace("%", "|");;
				rowData[2] = iObj.properties().getProperty("SI_KIND").toString();
				rowData[3] = iObj.properties().getProperty("SI_ID").toString();
				rowData[4] = iObj.properties().getProperty("SI_OWNER").toString();
				if (iObj.properties().getProperty("SI_CREATION_TIME").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_CREATION_TIME").getValue()).getTime());
					rowData[5] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));					
				} else {
					rowData[5] = "";
				}
				if (iObj.properties().getProperty("SI_UPDATE_TS").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_UPDATE_TS").getValue()).getTime());
					rowData[6] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));				
				} else {
					rowData[6] = "";
				}
				//Not bothered about universes
				rowData[7] = "";
				rowData[8] = "";
				if (iObj.properties().getProperty("SI_INSTANCE").getValue() != null) {
					rowData[9] = iObj.properties().getProperty("SI_INSTANCE").getValue().toString();
				} else {
					rowData[9] = "";
				}
				rowData[10] = arrStatus[(int)iObj.properties().getProperty("SI_SCHEDULE_STATUS").getValue()];
				
				wtExcel.writeSheet(curFolder, rowData);
			}
		} catch (Exception e) {
			System.out.println(sPath + " -- " + iObj.properties().getProperty("SI_NAME").toString());
			System.out.println("GETINSTANCES  --  " + e.toString());
		}

	}
	*/
	
/*
private static void getUserReports(String sSheet, String sUserID, int sParID) {
		
		String[] rowData = new String[12];	
		IInfoObject iObj = null;
		IInfoObjects subFolders = null;
		IWebi iRep = null;
		String sUnvs = "";
		Object[] oUnv = null;
			
		String query = "Select top " + iLimit +  " * from ci_infoobjects where si_parentid = " + sParID;
		try {
			subFolders = iStore.query(query);
			System.out.println("Getting all " + subFolders.size() + " " + sSheet + " reports for " + sParID);
			for (int j=0;j < subFolders.size(); j++) {
				iObj = (IInfoObject)subFolders.get(j);
				rowData[0] = sUserID;
				rowData[11] = iObj.properties().getProperty("SI_PARENTID").toString();
				rowData[1] = iObj.properties().getProperty("SI_NAME").toString().replace("%26", "&").replace("%", "|");
				rowData[2] = iObj.properties().getProperty("SI_KIND").toString();
				rowData[3] = iObj.properties().getProperty("SI_ID").toString();
				rowData[4] = iObj.properties().getProperty("SI_OWNER").toString();
				if (iObj.properties().getProperty("SI_CREATION_TIME").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_CREATION_TIME").getValue()).getTime());
					rowData[5] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));					
				} else {
					rowData[5] = "";
				}
				if (iObj.properties().getProperty("SI_UPDATE_TS").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_UPDATE_TS").getValue()).getTime());
					rowData[6] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));				
				} else {
					rowData[6] = "";
				}
				if (iObj.properties().getProperty("SI_INSTANCE").getValue() != null) {
					rowData[9] = iObj.properties().getProperty("SI_INSTANCE").getValue().toString();
				} else {
					rowData[9] = "";
				}
				if (rowData[9].equals("true")) {
					//This is an instance
					rowData[10] = arrStatus[(int)iObj.properties().getProperty("SI_SCHEDULE_STATUS").getValue()];
				} else {
					rowData[10] = "";
				}
				
				if (rowData[2].equals("Webi")) {
					iRep = (IWebi)iObj;
					sUnvs = "";
					oUnv = iRep.getUniverses().toArray();
					System.out.print("Getting list of " + oUnv.length + " universes ... ") ;
					for (int z=0;z < oUnv.length;z++) {
						if (sUnvs.equals("")){
							sUnvs = oUnv[z].toString();
						} else {
							sUnvs = sUnvs + "^" + oUnv[z].toString();
						}
						System.out.print((z + 1) + " ");
					}
					rowData[7] = sUnvs;
					System.out.println(".");
					System.out.println("All Universes documented");
					//now get multi source universes
					sUnvs = "";
					oUnv = iRep.getDSLUniverses().toArray();
					System.out.print("Getting list of " + oUnv.length + " multi source universes ... ") ;
					for (int z=0;z < oUnv.length;z++) {
						if (sUnvs.equals("")){
							sUnvs = oUnv[z].toString();
						} else {
							sUnvs = sUnvs + "^" + oUnv[z].toString();
						}
						System.out.print((z + 1) + " ");
					}
					rowData[8] = sUnvs;
					System.out.println(".");
					System.out.println("All multisource Universes documented");
				} else {
					rowData[7] = "";
					rowData[8] = "";
				}			
				
				wtExcel.writeSheet(sSheet, rowData);
			}
		} catch (Exception e) {
			System.out.println(" -- " + iObj.properties().getProperty("SI_NAME").toString());
			System.out.println("GETUSERREPORTS  --  " + e.toString());
		}
	}
*/
	
private static void getAllReports() {
		
		String[] rowData = new String[14];	
		IInfoObject iObj = null;
		IInfoObject iObj2 = null;
		IWebi iRep = null;
		String sUnvs = "";
		Object[] oUnv = null;
		String query = "";
		String query2 = "";
		IInfoObjects iParent = null;
		String sFull = "";
		IFolder iFol = null;
		IInfoObjects subFolders = null;
		int iMaxID = 0;
			
		try {
			wtExcel.createSheet("Reports");
			rowData[0] = "SI_PARENTID";
			rowData[1] = "Name";
			rowData[2] = "SI_KIND";
			rowData[3] = "ID";
			rowData[4] = "Owner";
			rowData[5] = "Created";
			rowData[6] = "Last_Updated";
			rowData[7] = "Webi_Universes";
			rowData[8] = "Webi_Multi_Source_Universes";
			rowData[9] = "Instance?";
			rowData[10] = "Schedule Status";
			rowData[11] = "Parent_Type";
			rowData[12] = "Folder_Path";
			rowData[13] = "BO_System";
			wtExcel.writeHeader("Reports", rowData);
			
			//Now loop processing 1,000 records at a time
			for (;;) {
				query = "Select top 1000 * " 
						+ "from ci_infoobjects where si_kind in ('CrystalReport','Webi','FullClient','MDAnalysis','LCMJob','Flash',"
						+ "'XL.XcelsiusEnterprise','QaaWS','Pdf','Excel','Word','Powerpoint','Rtf','Txt','Shortcut','AFDashboardPage',"
						+ "'Analytic','Hyperlink','Publication','Xcelsius')"
						+" and si_id > " + iMaxID + " order by SI_ID ASC";
				subFolders = iStore.query(query);
				if (subFolders.size() == 0) {
					//Finished
					break;
				}
				System.out.println("Getting " + subFolders.size() + " reports (> " + iMaxID + ")");
				System.out.println("Using " + query);
				for (int j=0;j < subFolders.size(); j++) {
					System.out.println("Report " + (j + 1) + " of " + subFolders.size());
					iObj = (IInfoObject)subFolders.get(j);
					rowData[0] = iObj.properties().getProperty("SI_PARENTID").toString();
					rowData[1] = iObj.properties().getProperty("SI_NAME").toString().replace("%26", "&").replace("%", "|");
					rowData[2] = iObj.properties().getProperty("SI_KIND").toString();
					rowData[3] = iObj.properties().getProperty("SI_ID").toString();
					rowData[4] = iObj.properties().getProperty("SI_OWNER").toString();
					if (iObj.properties().getProperty("SI_CREATION_TIME").getValue() != null) {
						dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_CREATION_TIME").getValue()).getTime());
						rowData[5] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));					
					} else {
						rowData[5] = "";
					}
					if (iObj.properties().getProperty("SI_UPDATE_TS").getValue() != null) {
						dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_UPDATE_TS").getValue()).getTime());
						rowData[6] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));				
					} else {
						rowData[6] = "";
					}
					if (iObj.properties().getProperty("SI_INSTANCE").getValue() != null) {
						rowData[9] = iObj.properties().getProperty("SI_INSTANCE").getValue().toString();
					} else {
						rowData[9] = "";
					}
					if (rowData[9].equals("true")) {
						//This is an instance
						rowData[10] = arrStatus[(int)iObj.properties().getProperty("SI_SCHEDULE_STATUS").getValue()];
					} else {
						rowData[10] = "";
					}
					rowData[13] = mp.strCMS;
					query2 = "select top 1 * from ci_infoobjects where si_id = " + iObj.properties().getProperty("SI_PARENTID").toString();
					iParent= iStore.query(query2);
					if (iParent.size() == 0) {
						rowData[11] = "Not Found";
						rowData[12] = "Not Found";
					} else {
						iObj2 = (IInfoObject)iParent.get(0);
						rowData[11] = iObj2.getKind();
						if(rowData[11].equals("FavoritesFolder")) {
							rowData[12] = "User Folders/" + rowData[4];
						} else if (rowData[11].equals("Inbox")) {
							rowData[12] = iObj2.properties().getProperty("SI_OWNERID").getValue().toString();
						} else if (rowData[11].equals("ObjectPackage")) {
							rowData[12] = "Object";
						} else if (rowData[11].equals("Folder")) {
							iFol = (IFolder)iObj2;
							sFull = "";
							if (iFol.getPath() != null) {
								String path[]=iFol.getPath();
								for(int fi=0;fi<path.length;fi++) {
									sFull = path[fi] + "/" + sFull;
								}
								sFull = sFull + iFol.getTitle();
								rowData[12] = sFull;
							} else {
								rowData[12] = iFol.getTitle();
							}
						} else {
							rowData[12] = "";
						}
					}
					if (rowData[2].equals("Webi")) {
						iRep = (IWebi)iObj;
						sUnvs = "";
						oUnv = iRep.getUniverses().toArray();
						System.out.print("Getting list of " + oUnv.length + " universes ... ") ;
						for (int z=0;z < oUnv.length;z++) {
							if (sUnvs.equals("")){
								sUnvs = oUnv[z].toString();
							} else {
								sUnvs = sUnvs + "^" + oUnv[z].toString();
							}
							System.out.print((z + 1) + " ");
						}
						rowData[7] = sUnvs;
						System.out.println(".");
						System.out.println("All Universes documented");
						rowData[8] = "";
						System.out.println(".");
						System.out.println("All multisource Universes documented");
					} else {
						rowData[7] = "";
						rowData[8] = "";
					}			
				
					wtExcel.writeSheet("Reports", rowData);
					iMaxID = iObj.getID();
				}
			}
		} catch (Exception e) {
			System.out.println(" -- " + iObj.properties().getProperty("SI_NAME").toString());
			System.out.println("GETALLREPORTS  --  " + e.toString());
		}
	}

private static void getAllConnections() {
	
	String[] rowData = new String[11];	
	IInfoObject iObj = null;
	IInfoObjects subFolders = null;
		
	String query = "Select top " + iLimit +  " * from ci_appobjects where si_kind in ('CommonConnection', 'DFS.ConnectorConfiguration','DataFederator.DataSource','CCIS.DataConnection')";
	try {
		wtExcel.createSheet("All Connections");
		rowData[0] = "SI_ID";
		rowData[1] = "Name";
		rowData[2] = "SI_KIND";
		rowData[3] = "Parent_ID";
		rowData[4] = "Owner";
		rowData[5] = "Created";
		rowData[6] = "Last_Updated";
		rowData[7] = "Description";
		rowData[8] = "Parent_Folder";
		rowData[9] = "Database";
		rowData[10] = "BO_System";
		wtExcel.writeHeader("All Connections", rowData);
		subFolders = iStore.query(query);
		System.out.println("Getting all " + subFolders.size() + " connections");
		for (int j=0;j < subFolders.size(); j++) {
			iObj = (IInfoObject)subFolders.get(j);
			System.out.print("ID  ");
			rowData[0] = iObj.properties().getProperty("SI_ID").toString();
			System.out.print("NAME  ");
			rowData[1] = iObj.properties().getProperty("SI_NAME").toString().replace("%26", "&").replace("%", "|");
			System.out.print("KIND  ");
			rowData[2] = iObj.properties().getProperty("SI_KIND").toString();
			System.out.print("PARENTID  ");
			if (iObj.properties().getProperty("SI_PARENTID") == null) {
				rowData[3] = "";
			} else {
				rowData[3] = iObj.properties().getProperty("SI_PARENTID").toString();
			}
			System.out.print("OWNER  ");
			rowData[4] = iObj.properties().getProperty("SI_OWNER").toString();
			System.out.print("Description  ");
			if (iObj.properties().getProperty("SI_DESCRIPTION") == null) {
				rowData[7] = "";
			} else {
				rowData[7] = iObj.properties().getProperty("SI_DESCRIPTION").toString();
			}
			System.out.print("PARENT FOLDER  ");
			if (iObj.properties().getProperty("SI_PARENT_FOLDER") == null) {
				rowData[8] = "";
			} else {			
				rowData[8] = iObj.properties().getProperty("SI_PARENT_FOLDER").toString();
			}
			System.out.print("DB  ");
			if (iObj.properties().getProperty("SI_CONNECTION_DATABASE") == null) {
				rowData[9] = "";
			} else {
				rowData[9] = iObj.properties().getProperty("SI_CONNECTION_DATABASE").toString();
			}
			rowData[10] = mp.strCMS;
			System.out.print("CREATE  ");
			if (iObj.properties().getProperty("SI_CREATION_TIME") != null) {
				dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_CREATION_TIME").getValue()).getTime());
				rowData[5] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));					
			} else {
				rowData[5] = "";
			}
			System.out.print("UPDATE  ");
			if (iObj.properties().getProperty("SI_UPDATE_TS") != null) {
				dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_UPDATE_TS").getValue()).getTime());
				rowData[6] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));				
			} else {
				rowData[6] = "";
			}			
			wtExcel.writeSheet("All Connections", rowData);
			System.out.println(" -- DONE ---");
		}
	} catch (Exception e) {
		System.out.println(" -- " + iObj.properties().getProperty("SI_ID").toString());
		System.out.println("GETALLCONNECTIONS  --  " + e.toString());
	}
}
	
/*
	private static void getReports(String iParID, String sPath) {
		
		String[] rowData = new String[11];	
		IInfoObject iObj = null;
		IWebi iRep = null;
		String sUnvs = "";
		Object[] oUnv = null;
		
		String query = "Select top " + iLimit +  " * from ci_infoobjects where si_kind in ('CrystalReport','Webi','FullClient','MDAnalysis','LCMJob','Flash',"
				+ "'XL.XcelsiusEnterprise','QaaWS','PDF','Excel','Word','Powerpoint','Rtf','Txt') and si_parentid = " + iParID;
		try {
			IInfoObjects subFolders = iStore.query(query);
			System.out.println("Getting " + subFolders.size() + " reports");
			for (int j=0;j < subFolders.size(); j++) {
				iObj = (IInfoObject)subFolders.get(j);
				System.out.println("report " + (j + 1) + " - " + iObj.properties().getProperty("SI_NAME").toString());
				rowData[0] = sPath;
				rowData[1] = iObj.properties().getProperty("SI_NAME").toString().replace("%26", "&").replace("%", "|");
				rowData[2] = iObj.properties().getProperty("SI_KIND").toString();
				rowData[3] = iObj.properties().getProperty("SI_ID").toString();
				rowData[4] = iObj.properties().getProperty("SI_OWNER").toString();
				if (iObj.properties().getProperty("SI_CREATION_TIME").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_CREATION_TIME").getValue()).getTime());
					rowData[5] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));					
				} else {
					rowData[5] = "";
				}
				if (iObj.properties().getProperty("SI_UPDATE_TS").getValue() != null) {
					dtLocal.setTimeInMillis(((java.util.Date)iObj.properties().getProperty("SI_UPDATE_TS").getValue()).getTime());
					rowData[6] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));				
				} else {
					rowData[6] = "";
				}
				if (iObj.properties().getProperty("SI_INSTANCE").getValue() != null) {
					rowData[9] = iObj.properties().getProperty("SI_INSTANCE").getValue().toString();
				} else {
					rowData[9] = "";
				}
				if (rowData[9].equals("true")) {
					//This is an instance
					rowData[10] = iObj.properties().getProperty("SI_SCHEDULE_STATUS").getValue().toString();
				} else {
					rowData[10] = "";
				}
				
				if (rowData[2].equals("Webi")) {
					iRep = (IWebi)iObj;
					sUnvs = "";
					oUnv = iRep.getUniverses().toArray();
					System.out.print("Getting list of " + oUnv.length + " universes ... ") ;
					for (int z=0;z < oUnv.length;z++) {
						if (sUnvs.equals("")){
							sUnvs = oUnv[z].toString();
						} else {
							sUnvs = sUnvs + "^" + oUnv[z].toString();
						}
						System.out.print((z + 1) + " ");
					}
					rowData[7] = sUnvs;
					System.out.println(".");
					System.out.println("All Universes documented");
					//now get multi source universes
					sUnvs = "";
					oUnv = iRep.getDSLUniverses().toArray();
					System.out.print("Getting list of " + oUnv.length + " multi source universes ... ") ;
					for (int z=0;z < oUnv.length;z++) {
						if (sUnvs.equals("")){
							sUnvs = oUnv[z].toString();
						} else {
							sUnvs = sUnvs + "^" + oUnv[z].toString();
						}
						System.out.print((z + 1) + " ");
					}
					rowData[8] = sUnvs;
					System.out.println(".");
					System.out.println("All multisource Universes documented");
				} else {
					rowData[7] = "";
					rowData[8] = "";
				}			
				
				wtExcel.writeSheet(curFolder, rowData);
				if ((int)iObj.properties().getProperty("SI_CHILDREN").getValue() > 0) {
					getInstances(iObj.properties().getProperty("SI_ID").toString() , sPath);
				}
			}
		} catch (Exception e) {
			System.out.println(sPath + " -- " + iObj.properties().getProperty("SI_NAME").toString());
			System.out.println("GETREPORTS  --  " + e.toString());
		}
	}	
	*/

	public static void main(String[] args) {

		IEnterpriseSession enterpriseSession = null;
		IInfoObjects iObjects = null;
		IInfoObject iObject = null;
		IInfoObjects iObjsU = null;
		IInfoObject iObjU = null;		
		IProperties iProps = null;
		IUniverse iUnv = null;
		String sGrps = "";
		String userID; 
		int useridint;
		Boolean bUnx = false;
		ISessionMgr sessionMgr = null;
		
		File file;;
		FileOutputStream fos;
		PrintStream ps;
		
		String sSQL = "";
		//String sFolderID = "";
		//String sFolders[] = new String[2];
		String sUsers[] = new String[7];
		String sUniverses[] = new String[10];
		String sUnvRep[] = new String[3];
		//String[] repData = new String[11];
		//String[] rowHdr = new String[13];
		String strErr = "";
		Integer iUsr = 1;
		
		getHeap();
		
		mp = new GetBILogon();
		int ii = mp.getUserName();
		
		if (ii == 0) {
			
			try {
				file = new File(mp.strCMS + ".txt");
		        if (file.exists()) {
		        	file.delete();
		        }
				fos = new FileOutputStream(file);
				ps = new PrintStream(fos);
				System.setOut(ps);
				
				currentTime();
				
				frame = new JFrame("IBM SAP BI Tool");
				frame.setSize(400, 200);
				frame.setLocation(300, 300);
				frame.add(new JLabel("Extracting Data Please Wait ..."));
				frame.setVisible(true);
				frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
				
				wtExcel = new WriteToExcel(mp.strCMS + "_Reports.xlsx");
				
				System.out.println("Connecting...");
				sessionMgr = CrystalEnterprise.getSessionMgr();
				//enterpriseSession = sessionMgr.logon("Administrator", "W3dnesday", "WIN-250D8MB2MOL", "secEnterprise");
				enterpriseSession = sessionMgr.logon(mp.strUser, mp.strPassword, mp.strCMS, "secEnterprise");
				System.out.println("Connected");
				currentTime();
				
				iStore = (IInfoStore)enterpriseSession.getService("", "InfoStore");
				//sSQL = "select top " + iLimit +  " si_id, si_name, si_parentid, si_ancestor from ci_infoobjects where si_kind = 'Folder' and si_parentid = 23";
				//iObjects = iStore.query(sSQL);
				//wtExcel.createSheet("Hierarchy");
				//sFolders[0] = "Path";
				//sFolders[1] = "ID";
				//wtExcel.writeHeader("Hierarchy", sFolders);
				//for (int i=0;i < iObjects.size(); i++)
            		//{
					//getHeap();
					//iObject = (IInfoObject)iObjects.get(i);
					//iProps = iObject.properties();
					//sFolderID = iProps.getProperty("SI_NAME").toString();  
					//wtExcel.createSheet(sFolderID);
					//repData[0] = "Path";
					//repData[1] = "Name";
					//repData[2] = "SI_KIND";
					//repData[3] = "ID";
					//repData[4] = "Owner";
					//repData[5] = "Created";
					//repData[6] = "Updated";
					//repData[7] = "Webi Universes";
					//repData[8] = "Webi Multi Source Universes";
					//repData[9] = "Instance?";
					//repData[10] = "Schedule Status";
					//wtExcel.writeHeader(sFolderID, repData);				
					//curFolder = iProps.getProperty("SI_NAME").toString();
					//System.out.println("Processing high level folder " + curFolder);
					//sFolders[0] = sFolderID;
					//sFolders[1] = iProps.getProperty("SI_ID").toString();
					//wtExcel.writeSheet("Hierarchy", sFolders);
					//getReports(iProps.getProperty("SI_ID").toString(), iProps.getProperty("SI_NAME").toString());					
					//getSubFolders(iProps.getProperty("SI_ID").toString(), iProps.getProperty("SI_NAME").toString());
            	//}
				getHeap();
				System.out.println("Getting all reports");
				currentTime();
				getAllReports();
				System.out.println("Got all reports");
				getHeap();
				currentTime();
				strErr = wtExcel.closeXLS();
				if (strErr.equals("")) {
					System.out.println("Reports XLSX closed successfully");
				} else {
					throw new Exception("Report XSLX not closed. " + strErr); 
				}
				
				//USERS
				wtExcel = new WriteToExcel(mp.strCMS + "_Users.xlsx");
				wtExcel.createSheet("Users");
				sUsers[0] = "ID";
				sUsers[1] = "CUID";
				sUsers[2] = "Name";
				sUsers[3] = "Created";
				sUsers[4] = "Last_Logon";
				sUsers[5] = "User_Groups";
				sUsers[6] = "BO_System";
				wtExcel.writeHeader("Users", sUsers);
				
				/* rowHdr[0] = "Folder Owner ID";
				rowHdr[11] = "SI_PARENTID";
				rowHdr[1] = "Name";
				rowHdr[2] = "SI_KIND";
				rowHdr[3] = "ID";
				rowHdr[4] = "Owner";
				rowHdr[5] = "Created";
				rowHdr[6] = "Last Updated";
				rowHdr[7] = "Webi Universes";
				rowHdr[8] = "Webi Multi Source Universes";
				rowHdr[9] = "Instance?";
				rowHdr[10] = "Schedule Status";
				wtExcel.createSheet("Inboxes");
				wtExcel.writeHeader("Inboxes", rowHdr);
				wtExcel.createSheet("Favourites");
				wtExcel.writeHeader("Favourites", rowHdr); */
				sSQL = "SELECT top " + iLimit +  " SI_ID FROM CI_SYSTEMOBJECTS WHERE SI_KIND='User'";
				iObjects = iStore.query(sSQL);
				System.out.println(iObjects.size() + " users found");
				if (iObjects.getResultSize() > 0) {
					Iterator i = iObjects.iterator();
					while (i.hasNext()) {
						getHeap();
						System.out.println("User " + iUsr + " of " + iObjects.size());
						sGrps = "";
						iObject = (IInfoObject) i.next(); 
						useridint = iObject.getID(); 
						userID = ((Integer) useridint).toString();
						iObjsU = iStore.query("Select TOP 1 * From CI_SYSTEMOBJECTS Where SI_ID=" + userID);
						if (iObjsU.size() == 0) {
							System.out.println(iObject.getID() + " user account does not exist"); 
						} else {
							iObjU = (IInfoObject)iObjsU.get(0);
							IProperties iObjectProps = iObjU.properties(); 
							IProperty iObjectProp = iObjectProps.getProperty("SI_KIND"); 
							String iObjectPropVal = iObjectProp.getValue().toString(); 							
							sUsers[0] = iObjectProps.getProperty("SI_ID").toString();  
							sUsers[1] = iObjectProps.getProperty("SI_CUID").toString();
							sUsers[2] = iObjectProps.getProperty("SI_NAME").toString();							
							dtLocal.setTimeInMillis(((java.util.Date)iObjectProps.getProperty("SI_CREATION_TIME").getValue()).getTime());
							sUsers[3] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));
							if (iObjectProps.getProperty("SI_LASTLOGONTIME") != null) {
								dtLocal.setTimeInMillis(((java.util.Date)iObjectProps.getProperty("SI_LASTLOGONTIME").getValue()).getTime());
								sUsers[4] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));
							} else {
								sUsers[4] = "";
							}
							if (iObjectPropVal.equals("User"))  {
								IUser boxiUser = (IUser)iObjU; 
								// Obtain the set of groups which the user belongs to. 
								Object[] memberGroups = boxiUser.getGroups().toArray();
								for (int itt = 0; itt < memberGroups.length; itt++) { 
									if (sGrps.equalsIgnoreCase("")) {
										sGrps = memberGroups[itt].toString();
									} else {
										sGrps = sGrps + "^" + memberGroups[itt].toString();
									}
									// Retrieve each group from the infoStore. 
									//IInfoObjects result = iStore.query ("SELECT top " + iLimit +  " * FROM CI_SYSTEMOBJECTS WHERE SI_ID=" + memberGroups[itt]);
									//if (result.size() == 0) {
									//	if (sGrps.equals("")) {
									//		sGrps = "Could not find group with ID = " + memberGroups[itt];
									//	} else {
									//		sGrps = sGrps + "^" + "Could not find group with ID = " + memberGroups[itt];
									//	}
									//} else {
									//	IInfoObject iObject2 = (IInfoObject) result.get(0); 
									//	if (sGrps.equals("")) {
									//		sGrps = iObject2.getTitle();
									//	} else {
									//		sGrps = sGrps + "^" + iObject2.getTitle();
									//	}
									//}
								}
								//getUserReports("Inboxes", iObjectProps.getProperty("SI_ID").toString(), boxiUser.getInboxID());
								//getUserReports("Favourites", iObjectProps.getProperty("SI_ID").toString(), boxiUser.getFavoritesFolderID());
							} else {
								sGrps = "No groups available"; 
							}
							sUsers[5] = sGrps;
							sUsers[6] = mp.strCMS;
							wtExcel.writeSheet("Users", sUsers);
						}
						iUsr = iUsr + 1;
					}
					strErr = wtExcel.closeXLS();
					if (strErr.equals("")) {
						System.out.println("Users XLSX closed successfully");
					} else {
						throw new Exception("Users XSLX not closed. " + strErr); 
					}
				}
				
				getHeap();
				currentTime();
				//UNIVERSES
				wtExcel = new WriteToExcel(mp.strCMS + "_Universes.xlsx");
				wtExcel.createSheet("Universes");
				sUniverses[0] = "CUID";
				sUniverses[1] = "Name";
				sUniverses[2] = "Last_Updated";
				sUniverses[3] = "Revision";
				sUniverses[4] = "Description";
				sUniverses[5] = "Owner";
				sUniverses[6] = "Data_Connections";
				sUniverses[7] = "ID";
				sUniverses[8] = "SI_KIND";
				sUniverses[9] = "BO_System";
				wtExcel.writeHeader("Universes", sUniverses);
				wtExcel.createSheet("Universe Reports");
				sUnvRep[0] = "Universe_ID";
				sUnvRep[1] = "Report_ID";
				sUnvRep[2] = "BO_System";
				wtExcel.writeHeader("Universe Reports", sUnvRep);
				sSQL = "Select top " + iLimit +  " * FROM CI_APPOBJECTS WHERE SI_KIND in ('Universe','DSL.MetaDataFile')";
				iObjects = iStore.query(sSQL);
				System.out.println("Found " + iObjects.size() + " universes");
				for (int i=0;i < iObjects.size(); i++) {
					iObject = (IInfoObject)iObjects.get(i);
					iProps = iObject.properties();
					if (iProps.getProperty("SI_KIND").toString().equals("Universe")) {
						bUnx = false;
					} else {
						bUnx = true;
					}
					System.out.print(iProps.getProperty("SI_KIND").toString() + "(" + iProps.getProperty("SI_ID").toString() + ") universe " + i + " ... ");
					System.out.println(" GOT PROPERTIES");
					sUniverses[0] = iProps.getProperty("SI_CUID").toString();  
					System.out.print("SI_CUID  ");
					sUniverses[1] = iProps.getProperty("SI_NAME").toString();
					System.out.print("SI_NAME  ");
					if (iProps.getProperty("SI_UPDATE_TS") == null) {
						sUniverses[2] = "";
						System.out.print("SI_UPDATE_TS IS NULL  ");
					} else {
						dtLocal.setTimeInMillis(((java.util.Date)iProps.getProperty("SI_UPDATE_TS").getValue()).getTime());
						sUniverses[2] = dateFormatter.format((new java.util.Date(dtLocal.getTimeInMillis())));
						System.out.print("SI_UPDATE_TS  ");
					}
					if (iProps.getProperty("SI_REVISIONNUM") == null) {
						sUniverses[3] = "";
					} else {
						sUniverses[3] = iProps.getProperty("SI_REVISIONNUM").toString();
					}
					System.out.print("SI_REVISIONNUM  ");
					sUniverses[4] = iProps.getProperty("SI_DESCRIPTION").toString();
					System.out.print("SI_DESCRIPTION  ");
					sUniverses[5] = iProps.getProperty("SI_OWNER").toString();
					sUniverses[7] = iProps.getProperty("SI_ID").toString();
					sUniverses[8] = iProps.getProperty("SI_KIND").toString();
					sUniverses[9] = mp.strCMS;
					System.out.println("SI_OWNER");
					System.out.println("About to get Connections");
					
					if (!bUnx) {
						getHeap();
						iUnv = (IUniverse)iObject;
						Object[] oUnv = iUnv.getDataConnections().toArray();
						sGrps = "";
						System.out.println(i + " has " + oUnv.length);
						for (int itt = 0; itt < oUnv.length; itt++) { 
							System.out.println("doing -- " + itt);
							if (sGrps.equals("")) {
								sGrps = oUnv[itt].toString();
							} else {
								sGrps = sGrps + "^" + oUnv[itt].toString();
							}
							/* IInfoObjects result = iStore.query ("SELECT top " + iLimit +  " * FROM CI_APPOBJECTS WHERE SI_ID=" + oUnv[itt]);
							if (result.size() == 0) {
								if (sGrps.equals("")) {
									sGrps = "Could not find connection with ID = " + oUnv[itt];
								} else {
									sGrps = sGrps + "^" + "Could not find connection with ID = " + oUnv[itt];
								}
							} else {
								IInfoObject iObject2 = (IInfoObject) result.get(0); 
								if (sGrps.equals("")) {
									sGrps = iObject2.getTitle() + "(" + oUnv[itt].toString() + ")";
								} else {
									sGrps = sGrps + "^" + iObject2.getTitle() + "(" + oUnv[itt].toString() + ")";
								}
							} */
							System.out.println("done " + itt);
						}
						sUniverses[6] = sGrps;
						System.out.println(i);
						wtExcel.writeSheet("Universes", sUniverses);
						System.out.println(i + " reports");
						//xxxxx xmust be here
					
						if (iUnv.getWebis() == null) {
							System.out.println("iUnv is NULL");
						} 
						Object[] oURep = iUnv.getWebis().toArray();
						if (oURep == null) {
							System.out.println("getWedis is NULL");
						} else {			
							sUnvRep[0] = sUniverses[7];
							sUnvRep[2] = mp.strCMS;
							System.out.println("Found " + oURep.length + " reports for universe");
							for (int itt = 0; itt < oURep.length; itt++) {
								System.out.print(itt + "....");
								sUnvRep[1] = "" + oURep[itt];
								wtExcel.writeSheet("Universe Reports", sUnvRep);
								System.out.println("WRITTEN");
							}
						}
					} else {
						if (iProps.getProperty("SI_SL_UNIVERSE_TO_CONNECTIONS") != null) {
							sUniverses[6] = iProps.getProperty("SI_SL_UNIVERSE_TO_CONNECTIONS").toString();
						} else {
							sUniverses[6] = "";
						}
						//sUniverses[6] = "Cannot get data connections for multi source universe at present";
						wtExcel.writeSheet("Universes", sUniverses);
					}
				}
				
				getHeap();
				currentTime();
				System.out.println("Connections");
				getAllConnections();
				getHeap();
				currentTime();
				
				strErr = wtExcel.closeXLS();
				if (strErr.equals("")) {
					System.out.println("Universes XLSX closed successfully");
				} else {
					throw new Exception("Universes XSLX not closed. " + strErr); 
				}
				
				wtExcel = new WriteToExcel(mp.strCMS + "_Others.xlsx");
				//listRootFolder("19","Users","ci_systemobjects");
				getHeap();
				currentTime();
				listRootFolder("20","User Groups","ci_systemobjects");
				getHeap();
				currentTime();
				//listRootFolder("18","User Folders","ci_infoobjects");
				getHeap();
				currentTime();
				//listRootFolder("48","Inboxes","ci_infoobjects");
				getHeap();
				currentTime();
				listRootFolder("45","Categories","ci_infoobjects");
				getHeap();
				currentTime();
				listRootFolder("47","Personal Categories","ci_infoobjects");
				getHeap();
				currentTime();
				listRootFolder("22","Calendars","ci_systemobjects");
				strErr = wtExcel.closeXLS();
				if (strErr.equals("")) {
					System.out.println("Others XLSX closed successfully");
				} else {
					throw new Exception("Others XSLX not closed. " + strErr); 
				}
				
				getHeap();
				msgbox("Completed", "Extract for " + mp.strCMS + " has completed.", JOptionPane.INFORMATION_MESSAGE);
				System.out.println("Finished OK!");
			} catch (Exception ex) {
				System.out.println("XI3 Fail " + sGrps + " " + "MAIN  --  " + ex.toString());
				strErr = wtExcel.closeXLS();
				if (strErr.equals("")) {
					System.out.println("Others XLSX closed successfully");
					msgbox("Failed", "XLSX ERROR AND " + ex.toString(), JOptionPane.ERROR_MESSAGE);
				} else {
					msgbox("Failed", ex.toString(), JOptionPane.ERROR_MESSAGE);					
				}
			} finally {
				if (enterpriseSession != null) {
					enterpriseSession.logoff(); 
					System.out.println("Logged off");
				}
					frame.dispose();
			}

		}
		
	}

}
