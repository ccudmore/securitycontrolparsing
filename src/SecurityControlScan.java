import java.awt.font.NumericShaper.Range;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;

public class SecurityControlScan {	
    final static String HEADING = "Heading";
    
	public static void main(String[] args) {
		FileInputStream fis = null;
		String fileName = null;
        boolean doDebug = false;
		boolean doComments = false;
		if (args.length < 1) {System.out.println("Missing file name."); return;}
		fileName = args[0];
		if (args.length == 2 && args[1].contains("debug")) {doDebug = true;}
		if (args.length == 2 && args[1].contains("comments")) {doComments = true;}
		if (args.length == 3 && args[1].contains("debug")&& args[2].contains("comments")) {doDebug = true; doComments = true;}
        if (doDebug) {
        	long heapSize = Runtime.getRuntime().maxMemory();
        	System.out.println("Maximum heap size is "+heapSize);
			//CC 17-03-2022
			//July 4, 2022 - CC and DS added (comment to csv) capability  (uncludes extraction of 
			//Comment ID, Author, Comment Text and the comment anchor/highlighted document text)
			    		
			System.out.println("Compile Date: 04-07-2022");
        }
		
		int[] levelValues = new int[] {0,0,0,0,0,0};
		try {
			// Read the input file and split it apart using Apache POI
			
			fis = new FileInputStream(fileName);
            XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
            List<XWPFParagraph> pl = xdoc.getParagraphs();			
			            
            String sectionName = "";
						
            Map<String, ArrayList<String>> controlsMap = new HashMap<String,ArrayList<String>> ();
			Map<String, ArrayList<String>> commentsMap = new HashMap<String,ArrayList<String>> ();
			
            // Examine each paragraph and look for headings and comments
			for (XWPFParagraph p : pl) {
				String style = p.getStyle();
				if (style != null && style.contains(HEADING)) {
					int headingLevel = Integer.parseInt(style.substring(style.indexOf(HEADING)+HEADING.length()))-1;
					// when the current level changes, reset all of the child levels
					for (int i=headingLevel+1; i<levelValues.length-1; i++) {levelValues[i]=0;} // reset the child levels to zero
					levelValues[headingLevel] ++;
					
					sectionName = formatSectionName(levelValues, "Section") + " - " + p.getText(); // Get the new heading level
					if (doDebug) {
						System.out.println("Processing style entry: "+style+". Calculated as :"+sectionName);
					}
				} 
				
				// Look for controls within the paragraph and put them in the table
				for (String control : getControls(p.getText())) {
					// CC - if the control hasn't been seen before, create a new array of sections for it
					if (controlsMap.containsKey(control) == false ) {
						controlsMap.put(control, new ArrayList<String> ());
					}
					controlsMap.get(control).add(sectionName);
				}
				
				if (doComments) {
					// now see if the paragraph has any comments
					// Locate the startRange
					//List<CTMarkupRange> markups = p.getCTP().getCommentRangeEndList();
					List<CTMarkupRange> markups = p.getCTP().getCommentRangeStartList();
					for (CTMarkupRange markup: markups) {
						StringBuilder comments = new StringBuilder();
						StringBuilder commentHighlight = new StringBuilder();
						// Get the highlighted comment anchor text and replace "~" with a *-  we use ~ as delimiter for output
						for (XWPFRun run : p.getRuns()) {
						commentHighlight.append(run.text().replace("~", "*"));
						}
						XWPFComment commentText = xdoc.getCommentByID(markup.getId().toString());
						comments.append("~" + sectionName + "~" + commentText.getAuthor() + "~" + commentText.getText().replace("\n", "").replace("\r", "").replace("~", "*") + "~" + commentHighlight.toString());
						commentsMap.put(commentText.getId(), new ArrayList<String> ());
						commentsMap.get(commentText.getId()).add(comments.toString());
						if (doDebug) {
							System.out.println("----" + commentText.getId() + comments.toString());
						}
					}
				}
			}
			// sort the controls 
			Map<String, ArrayList<String>> sortedControls = sortControls(controlsMap);
			
			// DS - sortedComments requires and update to sort by CommentID as a Int instead of a string
			if (doComments) {
				Map<String, ArrayList<String>> sortedComments = sortControls(commentsMap);
				writeList(sortedComments, fileName+"-comments.csv");
			}
	    //	printList(sortedControls);
			writeList(sortedControls, fileName+"-controls.csv");
			
			
		} catch (Exception e) {
			System.out.println("ex"+e.getMessage());
			e.printStackTrace();
		}
	}
	
	private static boolean isNumeric(String s) {
		if (s == null) return false;
		try {double d = Double.parseDouble(s);}
		catch (NumberFormatException e) {return false;}
		return true;
	}
	
	private static Map<String, ArrayList<String>> sortControls(Map<String, ArrayList<String>> in) {
		Comparator<String> sortByControlId = new Comparator<String> () {
			@Override public int compare(String s1, String s2) {
				int base1 = s1.indexOf("(");
				int base2 = s2.indexOf("(");
				if (base1 == -1 && base2 == -1) return s1.compareTo(s2); // no enhancement, let them be
				// otherwise get the base control ids
				String sb1 = (base1 == -1)?s1:s1.substring(0,base1);
				String sb2 = (base2 == -1)?s2:s2.substring(0,base2);
				// if the base control ids are different, then no special sorting required
				if (!sb1.equals(sb2)) return s1.compareTo(s2);
				
				// now we have the same base control but different enhancements. Sort the custom enhancement order
				// get the enhancement
				String enh1 = s1.substring(base1+1, s1.length()-1);
				String enh2 = s2.substring(base2+1, s2.length()-1);
				
				// Sort so characters go before numbers
				// If both are numeric or both are non-numeric, no special sorting
				if ((isNumeric(enh1) && isNumeric(enh2)) || (!isNumeric(enh1) && !isNumeric(enh2)))
					return s1.compareTo(s2);
				// we now know the two enhancements are different types, sort accordingly
				return isNumeric(enh1)?+1:-1;
			}
		};
		Map<String, ArrayList<String>> sortedControls = new TreeMap<String, ArrayList<String>> (sortByControlId);
        for (Entry<String, ArrayList<String>> entry : in.entrySet()) {
        	sortedControls.put(entry.getKey(), entry.getValue());
        }
        return sortedControls;
	}
	
	// Format the document section name. This is needed because Word doesn't store
	// the actual section numbers, just section hierarchy. The section numbers need
	// to be re-created.
	private static String formatSectionName(int[] levelValues, String ident) {
		StringBuilder str = new StringBuilder();
		if (ident.length() > 0) str.append(ident).append(" ");
		for(int l: levelValues) {
			if (l == 0) break;
			str.append(l).append(".");
		}
		
		str.setLength(str.length()-1);
		return str.toString();
	}
	
	// Look for instances of the << and >> tags and create a list of all identified controls
	private static ArrayList<String> getControls(String str) throws Exception {
		ArrayList<String> controls = new ArrayList<String> ();
		
		if (str.length() > 0) {
			String[] strSplit = str.split("<<");
			for (int i=1; i< strSplit.length; i++) {
				int endHash = strSplit[i].indexOf(">>");
				if (endHash == -1) {throw new Exception("Unterminated Security Control tag in "+strSplit[i]);}
				controls.add(strSplit[i].substring(0, endHash));
			}
		}
		return controls;
	}

	public static <K, V> void printList(Map<K, V> map) {
        for (Map.Entry<K, V> entry : map.entrySet()) {
        	for (String section : (ArrayList<String>)entry.getValue()) {
                System.out.println("Key : " + entry.getKey() 
    				+ " Value : " + section);
        	}
        }
    }	

	public static <K, V> void writeList(Map<K, V> map, String fileName) throws Exception {
		FileWriter fw = new FileWriter(fileName);
        for (Map.Entry<K, V> entry : map.entrySet()) {
        	for (String section : (ArrayList<String>)entry.getValue()) {
            	fw.write("\""+entry.getKey()+"\",\""+section+"\"\r\n");
        	}
        }
        fw.close();
        System.out.println("Wrote "+map.size()+" entries to "+fileName+".");
    }	
	

}
