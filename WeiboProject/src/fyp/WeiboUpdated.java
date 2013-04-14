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
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Dictionary;
import java.util.Formatter;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.Number;
import jxl.write.Label;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.log4j.Logger;
import org.scribe.builder.ServiceBuilder;
import org.scribe.builder.api.SinaWeiboApi20;
import org.scribe.model.Token;
import org.scribe.model.Verifier;
import org.scribe.oauth.OAuthService;

import au.com.bytecode.opencsv.CSVWriter;

import org.joda.time.DateTime;
import org.joda.time.Months;

import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.gargoylesoftware.htmlunit.FailingHttpStatusCodeException;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlAnchor;
import com.gargoylesoftware.htmlunit.html.HtmlButton;
import com.gargoylesoftware.htmlunit.html.HtmlForm;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import com.gargoylesoftware.htmlunit.html.HtmlPasswordInput;
import com.gargoylesoftware.htmlunit.html.HtmlTextInput;

public class WeiboUpdated {
	// private static final String NETWORK_NAME = "SinaWeibo";
	// private static final String PROTECTED_RESOURCE_URL =
	// "https://api.weibo.com/2/account/get_uid.json";
	private static final Token EMPTY_TOKEN = null;
	private static final int MAX_STATUS_CNT = 50;
	private static final int ALL_DATA = 0;
	private static final int FEATURE_TYPE = 1;
	private static final int SCREEN_NAME_COL_IDX = 0;
	private static final int ACCOUNT_COL_IDX = 0;
	private static final int PASSWORD_COL_IDX = 1;
	private static final boolean KEYWORK_FILE_HAS_HEADER = true;
	private static final boolean NAME_LIST_FILE_HAS_HEADER = true;
	private static final boolean ACCOUNT_LIST_FILE_HAS_HEADER = true;
	private static final String ACCOUNT_LIST_FILE_NAME = "accounts.xls";
	private static final String KEYWORD_LIST_FILE_NAME = "keywords.xls";
	private static final String NAME_LIST_FILE_NAME = "fypnamelist.xls";
	private static final String OUTPUT_FILE_NAME = "keywordStat.xls";
	private static final int PERIOD_OF_MONTHS = 1;
	private static final String apiKey = "3858201645";
	private static final String apiSecret = "8863e7c72a9e39d569c572380d041bab";
	private static final String CALL_BACK_URL = "http://143.89.20.216";
	
	private static URI nameListPath = null;
	private static URI keywordListPath = null;
	private static List<String> accounts = null;
	private static List<String> keywordList = null;
	private static List<String> screenNameList = null;
	private static Map<String, String> acctMap = null;
	private static Map<String, Integer> keywordStat = null;
	private static Iterator<String> iAccounts = null;
	private static String currentAccount = null;
	
	private static org.apache.log4j.Logger log = Logger
			.getLogger(WeiboUpdated.class);

	public static void main(final String[] info) {
		// Name List Path
		init();
		// for(String keyword: keywordList)
		// log.debug(keyword);
		// for(String screenName: screenNameList)
		// log.debug(screenName);
		Timeline timeline = new Timeline();
		timeline.setToken(renewToken().getToken());

		Iterator<String> iScreenName = screenNameList.iterator();
		while (iScreenName.hasNext()) {
			String screenName = iScreenName.next();
			boolean hasNextPage = true;
			boolean outOfPeriod = false;
			int currentPageNum = 1;
			StatusWapper statusWapper = getStatusWapper(timeline,
					screenName, currentPageNum);
			do {
				Iterator<String> iKeyword = keywordList.iterator();

				for (Status status : statusWapper.getStatuses()) {
					// Get the Last #Months Status
					Months months = calMthsBetween(status);
					if (months.getMonths() > PERIOD_OF_MONTHS) {
						outOfPeriod = true;
						// Out of period.
						break;
					}
					while (iKeyword.hasNext()) {
						String keyword = iKeyword.next();
						Pattern pattern = Pattern.compile(Pattern
								.quote(keyword));
						Matcher matcher = pattern.matcher(status.getText());
						// log.info(String.format("Screen Name[%s];Status[%s];Keyword[%s]",
						// screen_name, status.getText(), keyword));
						if (matcher.find()) {
							Integer freq = keywordStat.get(keyword);
							log.info(String
									.format("Matched Keyword[%s], Current Frequency[%d], Screen Name[%s], Post[%s]",
											keyword, freq, screenName,
											status.getText()));
							keywordStat.put(keyword, (freq == null) ? 1
									: freq + 1);
						}
					}
				}

				// Check if there is remaining posts
				hasNextPage = (statusWapper.getTotalNumber() - (currentPageNum * MAX_STATUS_CNT)) > 0;
				if (hasNextPage && !outOfPeriod) {
					currentPageNum++;
					getStatusWapper(timeline, screenName,
							currentPageNum);
				}
			} while (hasNextPage && !outOfPeriod);
		}
		// Done
		outputResult(nameListPath.getPath());
	}

	private static void init() {
		nameListPath = getRes(NAME_LIST_FILE_NAME);
		keywordListPath = getRes(KEYWORD_LIST_FILE_NAME);
		keywordStat = new HashMap<String, Integer>();

		keywordList = readKeywords(new File(keywordListPath), 0,
				KEYWORK_FILE_HAS_HEADER);
		screenNameList = readScreenNames(new File(nameListPath), 0,
				SCREEN_NAME_COL_IDX, NAME_LIST_FILE_HAS_HEADER);
		initAccounts();
	}

	private static StatusWapper getStatusWapper(Timeline timeline,
			String screenName, int currentPageNum) {
		StatusWapper statusWapper = null;
		try {
			statusWapper = timeline.getUserTimelineByName(screenName,
					new Paging(currentPageNum, MAX_STATUS_CNT), ALL_DATA, FEATURE_TYPE);
		} catch (WeiboException e) {
			if(e.getErrorCode() == 10023){
				timeline.setToken(renewToken().getToken());
				try {
					statusWapper = timeline.getUserTimelineByName(screenName,
							new Paging(currentPageNum, MAX_STATUS_CNT), ALL_DATA, FEATURE_TYPE);
				} catch (WeiboException e1) {
					String msg = String.format("Error[%s],Account[%s]", e1.getError(), currentAccount);
					// Write result
					outputResult(nameListPath.getPath());
					log.error(msg, e1);
				}
			}
		}
		return statusWapper;
	}
	
	private static void outputResult(String path) {
		// Result
		try {
			String outputfilePath = path.replaceAll(NAME_LIST_FILE_NAME, OUTPUT_FILE_NAME);
			log.info(String.format("Output Result to File:%s", outputfilePath));
			File outputfile = new File(outputfilePath);
			if(outputfile.exists()){
				// Replace with a new file
				outputfile.delete();
			}
			WritableWorkbook outputWB = Workbook.createWorkbook(outputfile);
			WritableSheet sheet = outputWB.createSheet("Stat", 0);
			Iterator<String> iKeyword = keywordList.iterator();
			int keywordCol = 0;
			int freqCol = 1;
			int row = 1;
			sheet.addCell(new Label(0, keywordCol, "Keyword"));
			sheet.addCell(new Label(0, freqCol, "Frequency"));
			while (iKeyword.hasNext()) {
				String keyword = iKeyword.next();
				Label keywordLb = new Label(keywordCol, row, keyword);
				Number freq = new Number(freqCol, row++,
						(keywordStat.get(keyword) == null) ? 0
								: keywordStat.get(keyword));
				sheet.addCell(keywordLb);
				sheet.addCell(freq);
			}
			outputWB.write();
			outputWB.close();
		} catch (WriteException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static URI getRes(String resFileName) {
		URL resPathURL = WeiboUpdated.class.getClassLoader().getResource(
				resFileName);
		URI resPath = null;
		if(resPathURL != null) {
			try {
				resPath = resPathURL.toURI();
			} catch (URISyntaxException e) {
				e.printStackTrace();
			}
		}
		return resPath;
	}
	
	private static void initAccounts(){
		accounts = new ArrayList<String>();
		acctMap = new HashMap<String, String>();
		
		Workbook accountsWB = null;
		try {
			accountsWB = Workbook.getWorkbook(new File(getRes(ACCOUNT_LIST_FILE_NAME)));
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = accountsWB.getSheet(0);
		int numOfRow = sheet.getRows();

		// Only Get String Content.
		for (int i = 0; i < numOfRow; i++) {
			if (ACCOUNT_LIST_FILE_HAS_HEADER && i == 0) {
				continue;
			}

			Cell accountCell = sheet.getCell(ACCOUNT_COL_IDX, i);
			Cell passwordCell = sheet.getCell(PASSWORD_COL_IDX, i);
			if (accountCell.getType() == CellType.LABEL &&
					passwordCell.getType() == CellType.LABEL) {
				LabelCell LCAccountCell = (LabelCell) accountCell;
				LabelCell LCPasswordCell = (LabelCell) passwordCell;
				String account = LCAccountCell.getString();
				accounts.add(account);
				acctMap.put(account, LCPasswordCell.getString());
			}

		}

		accountsWB.close();
		
		iAccounts = accounts.iterator();
	}
	
	private static Token renewToken(){
		Token accessToken = null;
		if(iAccounts.hasNext()){
			currentAccount = iAccounts.next();
			String password = acctMap.get(currentAccount);
			try {
				accessToken = getToken(apiKey, apiSecret, currentAccount, password);
			} catch (FailingHttpStatusCodeException e) {
				e.printStackTrace();
			}
			if(!iAccounts.hasNext()){
				// Renew account iterator
				iAccounts = accounts.iterator();
			}
		}
		
		return accessToken;
	}

	private static Months calMthsBetween(final Status status) {
		DateTime createdDate = new DateTime(status.getCreatedAt());
		DateTime now = new DateTime();
		Months months = Months.monthsBetween(createdDate, now);
		return months;
	}

	public static List<String> readKeywords(final File file,
			final int sheetIdx, final boolean hasHeader) {
		List<String> keywords = new ArrayList<String>();
		Workbook keywordWB = null;
		try {
			keywordWB = Workbook.getWorkbook(file);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = keywordWB.getSheet(sheetIdx);
		int numOfCol = sheet.getColumns();
		int numOfRow = sheet.getRows();

		// Only Get String Content.
		for (int i = 0; i < numOfRow; i++) {
			if (hasHeader && i == 0) {
				continue;
			}
			for (int j = 0; j < numOfCol; j++) {
				Cell cell = sheet.getCell(j, i);
				if (cell != null && cell.getType() == CellType.LABEL) {
					LabelCell lc = (LabelCell) cell;
					keywords.add(lc.getString());
				}
			}
		}

		keywordWB.close();
		return keywords;
	}

	public static List<String> readScreenNames(final File file,
			final int sheetIdx, final int screenNameColIdx,
			final boolean hasHeader) {
		List<String> screenNames = new ArrayList<String>();
		Workbook screenNameWB = null;
		try {
			screenNameWB = Workbook.getWorkbook(file);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = screenNameWB.getSheet(sheetIdx);

		int numOfRow = sheet.getRows();

		// Only Get String Content.
		for (int i = 0; i < numOfRow; i++) {
			if (hasHeader && i == 0) {
				continue;
			}
			Cell cell = sheet.getCell(screenNameColIdx, i);
			if (cell != null && cell.getType() == CellType.LABEL) {
				LabelCell lc = (LabelCell) cell;
				screenNames.add(lc.getString());
			}
		}

		screenNameWB.close();
		return screenNames;
	}

	public static Token getToken(final String apiKey, final String apiSecret,
			final String email, final String password) {

		OAuthService service = new ServiceBuilder()
				.provider(SinaWeiboApi20.class).apiKey(apiKey)
				.apiSecret(apiSecret).callback(CALL_BACK_URL).build();
		// Scanner in = new Scanner(System.in);

		String authorizationUrl = service.getAuthorizationUrl(EMPTY_TOKEN);
		log.debug(authorizationUrl);

		String access_code = null;
		try{
			access_code = getAccessCode(email, password, authorizationUrl);
		}catch(FailingHttpStatusCodeException e){
			outputResult(nameListPath.getPath());
		}
		// log.info(access_code);
		Verifier verifier = new Verifier(access_code);

		return service.getAccessToken(EMPTY_TOKEN, verifier);
	}

	public static String getAccessCode(final String email,
			final String password, final String authorizationUrl) throws FailingHttpStatusCodeException{
		String accessCode = null;
		WebClient webClient = new WebClient(BrowserVersion.FIREFOX_17);
		HtmlPage page = null;
		try {
			page = webClient.getPage(authorizationUrl);
		} catch (FailingHttpStatusCodeException e) {
			e.printStackTrace();
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
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
		HtmlPage result = null;
		try {
			anchor.click();
			result = submitButton.click();
		} catch (IOException e) {
			e.printStackTrace();
		}
		int statusCode = result.getWebResponse().getStatusCode();
		if (statusCode == 200) {
			accessCode = result.getUrl().getQuery();
			if(accessCode != null){
				accessCode = accessCode.substring(5);
			}else{
				outputResult(nameListPath.getPath());
				return null;
			}
		} else {
			throw new FailingHttpStatusCodeException(result.getWebResponse());
		}

		webClient.closeAllWindows();

		return accessCode;
	}
}
