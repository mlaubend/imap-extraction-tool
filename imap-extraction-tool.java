package iMAP;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.io.Console;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.mail.*;
import org.apache.commons.io.IOUtils;

/***
 * Created by Mark Laubender
 * 
 ***/

public class DVExtract {
	private String username = null;
	private char[] pass;
	private int numEmails;
	java.util.Properties props = null;
	
	public DVExtract(){
		//configure jvm to use jsse security
		java.security.Security.addProvider(new com.sun.net.ssl.internal.ssl.Provider());
		//create properties for the session
		props = new java.util.Properties();
		//set the session up to use SSL for IMAP connections
		props.setProperty("mail.imap.socketFactory.fallback", "javax.net.ssl.SSLSocketFactory");
		//don't fallback to normal IMAP connections on failure
		props.setProperty("mail.imap.socketFactory.fallback", "false");
		//use simap port for imap/ssl connections
		props.setProperty("mail.imap.socketFactory.port", "993");
		//needed for maillistener
		props.setProperty("mail.imap.usesocketchannels", "true");
	}
				
	private void connectToMailServer(){
		try{
			//create session
			javax.mail.Session session = javax.mail.Session.getInstance(props, null);
			//create the mail store to store messages
			Store store = session.getStore("imaps");
			//connect
			store.connect("owamail.com", username, new String(pass));
			pass = null;
			System.out.println("connected to owamail");
		
			//open the inbox folder
			Folder DVFolder = store.getFolder(folder);
			DVFolder.open(Folder.READ_ONLY);		
			Message[] messages = DVFolder.getMessages();	
			
			if (numEmails > 1) 
				messages = sortMessages(messages);			
			else messages = new Message[]{messages[messages.length-1]};
			writeCSV(messages);
			
			DVFolder.close(false);
			store.close();
			System.out.println("Done!");
			
		}
		catch(MessagingException me){
			System.err.println("Incorrect Credentials"); //probably this
			me.printStackTrace();
		}
	}
	
	private void sortSeparatedMessageArrays(ArrayList<Message> messages){
		if (messages.size() < 1)
			return;
		
		try{
			for (int i = 0; i < messages.size(); i++){
				Message largest = messages.get(i);
				int pos = i;
				
				for (int j = i; j < messages.size(); j++){
					if (Integer.parseInt(messages.get(j).getSubject().split("#")[1]) > Integer.parseInt(largest.getSubject().split("#")[1])){
						largest = messages.get(j);
						pos = j;
					}
				}
				Message message = messages.get(i);
				messages.set(i, largest);
				messages.set(pos, message);
			}
		} catch (MessagingException me){
			me.printStackTrace();
		}
	}
	
	private Message[] sortMessages(Message[] messages){
		Message[] ret = new Message[numEmails];
		
		ArrayList<Message> list1 = new ArrayList<Message>();
		ArrayList<Message> list2 = new ArrayList<Message>();
		
		try{
			//separate list1 and list2 emails into separate lists 
			for (int i = messages.length-1; i > messages.length - numEmails-1; i--){
				if (messages[i].getSubject().contains(title)){
					list1.add(messages[i]);
				}
				else {
					list2.add(messages[i]);
				}
			}
			sortSeparatedMessageArrays(list1);
			sortSeparatedMessageArrays(list2);
			
			//put both arrays back together
			int i = 0;
			Boolean flip = true;
			int list1Position = 0;
			int list2Position = 0;
			while (i < numEmails){

				if (flip){
					if (list2Position > list2.size()){
						flip = !flip;
						continue;
					}
					ret[i] = list2.get(list2Position);
					list2Position++;
				}
				else {
					if (list1Position > list1.size()){
						flip = !flip;
						continue;
					}
					ret[i] = list1.get(list1Position);
					list1Position++;
				}
				flip = !flip;
				i++;
			}
		} catch (MessagingException me){
			me.printStackTrace();
		}
		return ret;
	}
	
	private void writeCSV(Message[] messages) {
		final int FILTER_ENTRY_MAX_SIZE = 15; //a single entry will never be more than 15 lines
		
		try{
			FileWriter writer = new FileWriter(parseDocumentTitle(messages) + ".csv");
			System.out.println("writing CSV");

			for (int i = 0; i < messages.length; i++){	
				writer.append(messages[i].getSubject() + ", Action\n");
				writer.append(prefix);
				writer.append("\n\n");
			
				//remove extra linebreaks, start only at ___ and end at ___, split on <br> tags
				//convert to byte array, then UTF-8 to get around encoding issues
				String content = "";
				if (messages[i].isMimeType("text/html")){	//TODO: add other mimetypes soon.  This WILL bite you in the ass later
					try{
						content = messages[i].getContent().toString();							
					} catch(UnsupportedEncodingException uex){
						InputStream is = messages[i].getInputStream();
						byte[] byteArray = IOUtils.toByteArray(is);
						content = new String(byteArray, "UTF-8");
					}
				}
				else {
					System.out.println("Let Mark know MimeType is not longer text/html"); //it's just easier this way...
				}
				String[] lines = ((String) content).replace("\n", "").replace("\r", "").split("<b id=\"2\">New:</b>")[1].split("<b id=\"3\">Modified")[0].split("<br>|<br/>|<br />"); 					


				for (int j = 0; j < lines.length; j++){ 
					String defaultAction = "";				
					Pattern filterTitleRegex = Pattern.compile("(&nbsp;){3} [0-9]{4,5}:"); //matches filter titles e.g. "&nbsp;&nbsp;&nbsp; 45193:"
					Matcher m = filterTitleRegex.matcher(lines[j]);
					
					if (m.find()){
						String[] filterTitle = parseFilterTitle(lines[j]);
						
						//EOF check for nested loop
						//when you get to the end, sometimes there is less than FILTER_ENTRY_MAX_SIZE lines
						int offset = FILTER_ENTRY_MAX_SIZE;
						if (lines.length - j < FILTER_ENTRY_MAX_SIZE) 
							offset = (lines.length - j) % FILTER_ENTRY_MAX_SIZE;
						
						for (int k = j+1; k < j+offset; k++){ //search within a single filter entry for default deployment		
							Pattern defaultDeployment = Pattern.compile("(.*)(Deployment:)(.*)[()]"); //matches default deployments e.g. "Deployment: Aggressive (Block / Notify)"
							m = defaultDeployment.matcher(lines[k]);

							if (m.find()){
								defaultAction = checkDefaultDeployment(lines[k]);
								break;
							}
							k++;
						}
						if (checkForRansomware(lines[j]))
							defaultAction = action;
						
						writer.append('\n' + filterTitle[0] + ": " + filterTitle[1] + ": " + filterTitle[2] + defaultAction);
					}					
				}
				writer.append("\n\n\n\n");
			}		
			writer.close();
		}
		catch(IOException ioe){
			System.err.println("IO exception: Error writing CSV, please try again");
			ioe.printStackTrace();
		}
		catch(MessagingException me){
			me.printStackTrace();
		}

	}
	
	private String parseDocumentTitle(Message[] messages){
		String list1 = "";
		String DV = "";
		Boolean list1Set = false;
		Boolean DVSet = false;
		try{
			for (int i = 0; i < messages.length; i++){ 
				String subject = messages[i].getSubject();
				
				if (subject.contains("list1") && !list1Set){
					list1 = subject;
					list1Set = true;
					continue;
				}
				else if (subject.contains(subject) && !DVSet){
					DV = subject;
					DVSet = true;
					continue;
				}
				
				if (subject.contains(subject))
					list1 += ", #" + subject.split("#")[1];
				else DV += ", #" + subject.split("#")[1];
			}
			
		} catch(MessagingException me){
			me.printStackTrace();
		}
		return DV + " " + list1;
	}
	
	private String[] parseFilterTitle(String string){
		String[] splitline = string.split(":");
		String[] stripNbsp = splitline[0].split(" ");
		String[] stripBr = splitline[2].split("<br");

		String[] ret = {stripNbsp[1], splitline[1], stripBr[0]};
		return ret;
	}
	
	private String checkDefaultDeployment(String string){
		String ret = "";
		
		Matcher m = Pattern.compile("\\(([^)]+)\\)").matcher(string); //pulls out anything inside parentheses
		while (m.find()){
			if (m.group(1).contains(default)){
				ret = action;
			}
		}
		return ret;
	}
	
	private Boolean checkForRansomware(String filter){
		ArrayList<String> stringsThatMatchRansomware = new ArrayList<String>();
		stringsThatMatchRansomware.add("crypt");
		stringsThatMatchRansomware.add("ransom");
		stringsThatMatchRansomware.add("locker");
		stringsThatMatchRansomware.add("locky");
		stringsThatMatchRansomware.add("krypt");
				
		for (String ransomware : stringsThatMatchRansomware){
			if (filter.toLowerCase().contains(ransomware)){
				return true;				
			}
		}
		return false;
	}
	
	public void setUsername(String username){
		this.username = username;
	}
	
	public void setPass(char[] pass){
		this.pass = pass;
	}
	
	public void setNumEmails(int numEmails){
		this.numEmails = numEmails;
	}
	
	public static void main(String[] args){
		DVExtract im = new DVExtract();
		Console console = System.console();

		System.out.print("Enter AD Username: ");
		Scanner scanner = new Scanner(System.in);
		im.setUsername(scanner.next());
		im.setPass(console.readPassword("Enter AD Password: "));
		scanner.nextLine();
		System.out.print("How many Emails?: ");
		im.setNumEmails(scanner.nextInt());
		scanner.close();

		im.connectToMailServer();
	}
}//end class

//adding something for commit
