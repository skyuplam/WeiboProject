package fyp;

import weibo4j.Timeline;
import weibo4j.model.Paging;
import weibo4j.model.Status;
import weibo4j.model.StatusWapper;
import weibo4j.model.WeiboException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.log4j.Logger;
import org.scribe.builder.ServiceBuilder;
import org.scribe.builder.api.SinaWeiboApi20;
import org.scribe.model.Token;
import org.scribe.model.Verifier;
import org.scribe.oauth.OAuthService;

import au.com.bytecode.opencsv.CSVWriter;

import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlAnchor;
import com.gargoylesoftware.htmlunit.html.HtmlButton;
import com.gargoylesoftware.htmlunit.html.HtmlForm;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import com.gargoylesoftware.htmlunit.html.HtmlPasswordInput;
import com.gargoylesoftware.htmlunit.html.HtmlTextInput;

public class WeiboUpdated {
	private static final String NETWORK_NAME = "SinaWeibo";
	private static final String PROTECTED_RESOURCE_URL = "https://api.weibo.com/2/account/get_uid.json";
	private static final Token EMPTY_TOKEN = null;
	private static org.apache.log4j.Logger log = Logger
            .getLogger(WeiboUpdated.class);
	public static void main(String[] info) throws WeiboException, IOException {
		String apiKey = "3858201645";
		String apiSecret = "8863e7c72a9e39d569c572380d041bab";
		String email = "wesleyhyfu@gmail.com";
		String password = "icemoon";
		URL nameListPath = WeiboUpdated.class.getClassLoader().getResource("fypnamelist.xls");
		if(nameListPath != null){
			log.debug("Name List Path:" + nameListPath.toString());
		}
		
		Timeline tm = new Timeline();
		Paging pag = new Paging();
		//
		String[] keyword = new String[] { "有", "你", "我" };
		int[] word1 = new int[20];
		int[] word2 = new int[20];
		int[] word3 = new int[20];
		String[][] data = new String[keyword.length][word1.length];
		String[][] count = null;

		Token accessToken = null;

		try {
			accessToken = get_token(apiKey, apiSecret, email, password);
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		String access_token = accessToken.getSecret();

		tm.client.setToken(access_token);
		try { // A
			int y1 = 0;
			int y2 = 0;
			int y3 = 0;
			
			Workbook workbook = Workbook.getWorkbook(new File(nameListPath.toURI()));
			Sheet sheet = workbook.getSheet(0);
			String[] name = new String[sheet.getRows()];
			for (int i = 0; i < name.length; i++) {// B
				int w1 = 0;
				int w2 = 0;
				int w3 = 0;
				int v = 0;
				name[i] = (sheet.getCell(0, i).getContents());
				String filePath1 = "D:\\FYP\\";
				String filePath2 = name[i];
				String filePath3 = ".csv";
				byte[] bom = { (byte) 0xFF, (byte) 0xFE };
				FileOutputStream fileOutputStream = new FileOutputStream(
						filePath1 + filePath2 + filePath3);
				fileOutputStream.write(bom);
				OutputStreamWriter outputStreamWriter = new OutputStreamWriter(
						fileOutputStream, "UTF-16LE");
				String last = name[i];
				Matcher m2 = Pattern.compile("Number").matcher(last);
				if (m2.find()) { // C
					CSVWriter writer = null;
					byte[] bom2 = { (byte) 0xFF, (byte) 0xFE };
					FileOutputStream fileOutputStream2 = new FileOutputStream(
							filePath1 + filePath2 + filePath3);
					fileOutputStream2.write(bom2);
					OutputStreamWriter outputStreamWriter2 = new OutputStreamWriter(
							fileOutputStream2, "UTF-16LE");
					writer = new CSVWriter(new OutputStreamWriter(
							fileOutputStream2, "UTF-16LE"), '\t');
					name = ArrayUtils.add(name, 0, " ");
					writer.writeNext(name);

					String a = Arrays.toString(word1);
					String sa[] = a.substring(1, a.length() - 1).split(",");
					String a2 = Arrays.toString(word2);
					String sa2[] = a2.substring(1, a2.length() - 1).split(",");
					String a3 = Arrays.toString(word3);
					String sa3[] = a3.substring(1, a3.length() - 1).split(",");
					String[] kw1 = new String[] { keyword[0] };
					String[] kw2 = new String[] { keyword[1] };
					String[] kw3 = new String[] { keyword[2] };
					kw1 = ArrayUtils.addAll(kw1, sa);
					kw2 = ArrayUtils.addAll(kw2, sa2);
					kw3 = ArrayUtils.addAll(kw3, sa3);

					data = new String[][] { kw1, kw2, kw3 };

					for (int op = 0; op < 3; op++) {
						writer.writeNext(data[op]);
					}
					writer.close();
					fileOutputStream.close();
					outputStreamWriter2.close();
					break;
				} // C
				pag.setPage(1);
				pag.setCount(100);
				StatusWapper status = tm.getUserTimelineByName(name[i], pag, 0,
						0);
				for (Status s : status.getStatuses()) { // D

					outputStreamWriter.write("\r\n");
					outputStreamWriter.write(s.toString());
					outputStreamWriter.write("\r\n");

					for (int q = 0; q < keyword.length; q++) {
						int k = 0;
						String words = s.getText();
						Matcher m = Pattern.compile(keyword[q]).matcher(words);
						if (m.find()) {
							k++;
							if (q == 0)
								++w1;
							if (q == 1)
								++w2;
							if (q == 2)
								++w3;
						}
						v = v + k;
						outputStreamWriter.write(keyword[q] + " appear for "
								+ k);
						outputStreamWriter.write("\r\n");
					}
				} // D
				word1[i] = w1;
				word2[i] = w2;
				word3[i] = w3;
				outputStreamWriter.write("Total number of 有: " + w1 + "\r\n");
				outputStreamWriter.write("Total number of 你: " + w2 + "\r\n");
				outputStreamWriter.write("Total number of 我: " + w3 + "\r\n");
				outputStreamWriter.write("Total number of the words: " + v
						+ "\r\n");
				y1 = y1 + w1;
				y2 = y2 + w2;
				y3 = y3 + w3;
				outputStreamWriter.close();
			}// B
			SimpleDateFormat format = new SimpleDateFormat(
					" EEE MMM dd HH:mm:ss zzz yyyy ");
			String time = " Wed Dec 16 00:00:00 CST 2012 ";
			Date date = null;
			date = format.parse(time);
			System.out.println("Format To times:");
			System.out.println(date.getTime());
			Timestamp ts = new Timestamp(date.getTime());
			System.out.println(ts);
		} // A
		catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Token get_token(String apiKey, String apiSecret,
			String email, String password) throws Exception {

		OAuthService service = new ServiceBuilder()
				.provider(SinaWeiboApi20.class).apiKey(apiKey)
				.apiSecret(apiSecret).callback("http://143.89.20.216").build();
		Scanner in = new Scanner(System.in);

		String authorizationUrl = service.getAuthorizationUrl(EMPTY_TOKEN);

		String access_code = get_access_code(email, password, authorizationUrl);
		// log.info(access_code);
		Verifier verifier = new Verifier(access_code);

		return service.getAccessToken(EMPTY_TOKEN, verifier);

	}

	public static String get_access_code(String email, String password,
			String authorizationUrl) throws Exception {
		String accessCode = null;
		WebClient webClient = new WebClient(BrowserVersion.FIREFOX_17);
		HtmlPage page = webClient.getPage(authorizationUrl);
		HtmlForm form = page.getFormByName("authZForm");
		// HtmlSubmitInput button = form.getInputByName("submitbutton");
		@SuppressWarnings("unchecked")
		List<HtmlAnchor> anchors = (List<HtmlAnchor>) page
				.getByXPath("//a[@class='WB_btn_login formbtn_01']");
		HtmlTextInput userId = form.getInputByName("userId");
		HtmlPasswordInput passWd = form.getInputByName("passwd");
		HtmlButton submitButton = (HtmlButton) page.createElement("button");
		submitButton.setAttribute("type", "submit");
		form.appendChild(submitButton);
		userId.setValueAttribute(email);
		passWd.setValueAttribute(password); // HTTP POST
		HtmlAnchor anchor = anchors.get(0);
		HtmlPage anchorResult = anchor.click();
		HtmlPage result = submitButton.click();
		int statusCode = result.getWebResponse().getStatusCode();
		if (statusCode == 200)
			accessCode = result.getUrl().getQuery();
		else {
			StringBuilder sb = new StringBuilder();
			accessCode = sb.append(statusCode).toString();
		}

		webClient.closeAllWindows();

		return accessCode.substring(5);
	}
}
