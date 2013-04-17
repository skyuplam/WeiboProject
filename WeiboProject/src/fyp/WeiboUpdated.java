package fyp;

import weibo4j.Timeline;
import weibo4j.model.Paging;
import weibo4j.model.Status;
import weibo4j.model.StatusWapper;
import weibo4j.model.WeiboException;

import java.io.Console;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.scribe.builder.ServiceBuilder;
import org.scribe.builder.api.SinaWeiboApi20;
import org.scribe.model.Token;
import org.scribe.model.Verifier;
import org.scribe.oauth.OAuthService;

import org.joda.time.DateTime;
import org.joda.time.Months;

import com.gargoylesoftware.htmlunit.FailingHttpStatusCodeException;
import com.google.gson.Gson;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;

public class WeiboUpdated {
	private static final String OUTPUT_FILE_KEYWORD_SHEET_NAME = "KeywordCounting";
	private static final String OUTPUT_FILE_DATE_KEYWORD_SHEET_NAME = "KeywordCountingByMth";
	private static final Token EMPTY_TOKEN = null;
	// Weibo Timeline, get the number of status per request. Default 50, MAX 200
	private static final int MAX_STATUS_CNT = 200;
	// Weibo Timeline, 0: All data (Default). 1: On device data
	private static final int ALL_DATA = 0;
	// Weibo Timeline, 1: get text
	private static final int FEATURE_TYPE = 1;
	// The column index for screen name
	private static final int SCREEN_NAME_COL_IDX = 0;
	// The column index for weibo account id
	private static final int ACCOUNT_COL_IDX = 0;
	// The column index for weibo account password
	private static final int PASSWORD_COL_IDX = 1;
	// Indicating that whether the keyword file has header
	private static final boolean KEYWORK_FILE_HAS_HEADER = true;
	// Indicating that whether the name line file has header
	private static final boolean NAME_LIST_FILE_HAS_HEADER = true;
	// Indicating that whether the account list file has header
	private static final boolean ACCOUNT_LIST_FILE_HAS_HEADER = true;
	// Please fill the account list file name here
	private static final String ACCOUNT_LIST_FILE_NAME = "accounts.xls";
	// Please fill the keyword list file name here
	private static final String KEYWORD_LIST_FILE_NAME = "keywords.xls";
	// Please fill the name list file name here
	private static final String NAME_LIST_FILE_NAME = "fypnamelist.xls";
	// Please fill the output file name here
	private static final String OUTPUT_FILE_NAME = "keywordStat.xls";
	private static final String OUTPUT_STATUS_FILE_NAME = "status.json";
	private static final String OUTPUT_STATUS_MODIFIED_FILE_NAME = "status_analysised.json";
	// The name of key column header for the output file
	private static final String OUTPUT_FILE_KEY_COL_NAME = "Keywords";
	// The name of screen name column header for the output file
	private static final String OUTPUT_FILE_SCREEN_NAME_COL_NAME = "Screen Name";
	// The name of Date column header for the output file
	private static final String OUTPUT_FILE_DATE_COL_NAME = "Date";
	// The name of value column header for the output file
	private static final String OUTPUT_FILE_VALUE_COL_NAME = "Keyword Frequency";
	// 
	private static final String OUTPUT_FILE_MENTIONED_COL_NAME = "Number of People Mentioned";
	// The key column index for the output file
	private static final int OUTPUT_FILE_KEY_COL_IDX = 0;
	// the value column index for the output file
	private static final int OUTPUT_FILE_VALUE_COL_IDX = 1;
	// The period of months for extracting the status from Weibo. For example,
	// Last 6 months
	private static final int PERIOD_OF_MONTHS = 6; // 6 Months
	// API Key from open.weibo.com
	private static final String apiKey = "3858201645";
	// API Secred key from open.weibo.com
	private static final String apiSecret = "8863e7c72a9e39d569c572380d041bab";
	// The URL used to redirect
	private static final String CALL_BACK_URL = "http://143.89.20.216";
	// Used for slow down the process
	private static final long PAUSE_PERIOD = 30000; // Wait 30 seconds
	private static final long PAUSE_PERIOD_PER_REQUEST = 2000; // Wait 2 seconds
																// before each
																// request
	// Start from last result if true. 
	// Set to True ONLY if the process is not finished last time.
	private static final boolean LOAD_PREVIOUS_RESULT = false;
	// Timeout for getting user input for the CAPTCHA
	private static final long TIME_OUT_SECONDS = 90;
	// Time to get inpii
	private static final long WAIT_FOR_INPUT_MILLIS = 3000;
	
	private static final int ACCESS_CODE_LENGTH = 32;
	// The page title of the redirected website. This will be used to stop the
	// waiting of CHATCHAS input. If you change the title of the redirected page,
	// you have to change this setting accordingly.
	private static final String PAGE_TITLE = "Index of /";
	private static final String VCODE_SECTION_XPATH = "//div[@class='oauth_login_form']/p[@node-type='validateBox']";
	private static final String SCREEN_NAME_PROP_KEY = "Current Screent Name";
	private static final String ACCOUNT_PROP_KEY = "Current Account";
	private static final String PAGE_NUM_PROP_KEY = "Current Page Number";
	private static final String FINISHED_FLAG_KEY = "Finished Flag";
	

	private static URI nameListPath = null;
	private static URI keywordListPath = null;
	private static List<String> accounts = null;
	private static List<String> keywordList = null;
	private static List<String> screenNameList = null;
	private static List<Post> statusesList = null;
	private static Map<String, String> acctMap = null;
	private static Map<String, List<Post>> keywordStat = null;
	private static Map<String, Map<String, Map<String, Integer>>> dateKeywordsNameMap = null;
	private static Iterator<String> iAccounts = null;
	private static String currentAccount = "";
	private static String currentScreenName = "";
	private static int currentPageNumber = 1;
	private static boolean finished = true;
	
	// Set to true when you need to collect Weibo data online
	private static boolean collectStatusOnline = false;
	

	private static org.apache.log4j.Logger log = Logger
			.getLogger(WeiboUpdated.class);

	public static void main(final String[] info) {
		// Name List Path
		init();
		// getToken(apiKey, apiSecret, "wesleyhyfu@gmail.com", "Abc123456");
		String input = "n";
		Console console = System.console();
		String collectStatusOnlinePrompt = "Are you going to collect status from weibo online? [Y/n] >";
		input = console.readLine(collectStatusOnlinePrompt);
		if(input.equals("Y")){
			collectStatusOnline = true;
		}else{
			collectStatusOnline = false;
		}
		// Start to Analysis
		if(collectStatusOnline){
			Timeline timeline = new Timeline();
			timeline.setToken(renewToken().getToken());
			collectWeiboStatus(timeline);
			analysisWeiboStatuses();
			saveStatuses();
		}else{
			readStatusFile();
			initKeywordStat();
		}
		stat();
		// Done, write the result to xls file
		outputResult();
	}

	private static void init() {
		nameListPath = getRes(NAME_LIST_FILE_NAME);
		keywordListPath = getRes(KEYWORD_LIST_FILE_NAME);
		keywordStat = new HashMap<String, List<Post>>();
		dateKeywordsNameMap = new HashMap<String, Map<String, Map<String, Integer>>>();
		keywordList = readKeywords(new File(keywordListPath), 0,
				KEYWORK_FILE_HAS_HEADER);
		screenNameList = readScreenNames(new File(nameListPath), 0,
				SCREEN_NAME_COL_IDX, NAME_LIST_FILE_HAS_HEADER);
		initAccounts();
		loadProperties();
//		readKeywordStat(true);
		// Fast-forward to the current account
		while (iAccounts.hasNext() && !currentAccount.isEmpty()) {
			String account = iAccounts.next();
			if (account.equals(currentAccount)) {
				break;
			}
		}
	}
	
	private static void initKeywordStat(){
		Iterator<Post> iStatus = statusesList.iterator();
		while(iStatus.hasNext()){
			Post post = iStatus.next();
			List<String> keywords = post.getKeywords();
			if(keywords != null && !keywords.isEmpty()){
				for(String keyword : keywords){
					List<Post> postList = keywordStat.get(keyword);
					if(postList == null){
						postList = new ArrayList<Post>();
					}
					postList.add(post);
					keywordStat.put(keyword, postList);
				}
			}
			
		}
	}
	
	private static void readStatusFile(){
		File file = new File(OUTPUT_STATUS_MODIFIED_FILE_NAME);
		if(!file.exists())
			return;
		try {
			FileInputStream fis = new FileInputStream(file);
			statusesList = loadWeiboStatuses(fis);
			fis.close();
		} catch (FileNotFoundException e) {
			log.error(e.getMessage(), e);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
	}
	
	private static List<Post> loadWeiboStatuses(InputStream in) {
		List<Post> posts = new ArrayList<Post>();
		try {
			JsonReader reader = new JsonReader(new InputStreamReader(in, "UTF-8"));
			Gson gson = new Gson();
			reader.beginArray();
			while(reader.hasNext()){
				Post post = gson.fromJson(reader, Post.class);
				posts.add(post);
			}
			reader.endArray();
			reader.close();
		} catch (UnsupportedEncodingException e) {
			log.error(e.getMessage(), e);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
		return posts;
	}

//	private static void readKeywordStat(boolean hasHeader) {
//		if(!LOAD_PREVIOUS_RESULT)
//			return ;
//		Workbook wb = null;
//		Sheet sheet = null;
//		try {
//			String filePath = getRes(NAME_LIST_FILE_NAME).getPath().replaceAll(
//					NAME_LIST_FILE_NAME, OUTPUT_FILE_NAME);
//			File file = new File(filePath);
//			if (!file.exists()) {
//				return;
//			}
//			wb = Workbook.getWorkbook(file);
//			sheet = wb.getSheet(0);
//		} catch (BiffException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		} catch (IndexOutOfBoundsException e) {
//			e.printStackTrace();
//		}
//
//		// Only Get String Content.
//		for (int i = 0; i < sheet.getRows(); i++) {
//			if (hasHeader && i == 0) {
//				continue;
//			}
//
//			Cell keyCell = sheet.getCell(OUTPUT_FILE_KEY_COL_IDX, i);
//			Cell freqCell = sheet.getCell(OUTPUT_FILE_VALUE_COL_IDX, i);
//			if (keyCell.getType() == CellType.LABEL
//					&& freqCell.getType() == CellType.NUMBER) {
//				LabelCell lc = (LabelCell) keyCell;
//				NumberCell freq = (NumberCell) freqCell;
//				keywordStat.put(lc.getString(), Math.round(freq.getValue()));
//			}
//
//		}
//
//		wb.close();
//	}

	private static void initAccounts() {
		accounts = new ArrayList<String>();
		acctMap = new HashMap<String, String>();

		Workbook accountsWB = null;
		Sheet sheet = null;
		try {
			accountsWB = Workbook.getWorkbook(new File(
					getRes(ACCOUNT_LIST_FILE_NAME)));
			sheet = accountsWB.getSheet(0);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}

		int numOfRow = sheet.getRows();

		// Only Get String Content.
		for (int i = 0; i < numOfRow; i++) {
			if (ACCOUNT_LIST_FILE_HAS_HEADER && i == 0) {
				continue;
			}

			Cell accountCell = sheet.getCell(ACCOUNT_COL_IDX, i);
			Cell passwordCell = sheet.getCell(PASSWORD_COL_IDX, i);
			if (accountCell.getType() == CellType.LABEL
					&& passwordCell.getType() == CellType.LABEL) {
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

	private static void loadProperties() {
		if(!LOAD_PREVIOUS_RESULT)
			return ;
		PropertiesAgent pa = new PropertiesAgent();
		if (pa.loadProperty(SCREEN_NAME_PROP_KEY, "") == null) {
			return;
		}
		currentScreenName = pa.loadProperty(SCREEN_NAME_PROP_KEY, "");
		currentPageNumber = Integer.parseInt(pa.loadProperty(PAGE_NUM_PROP_KEY,
				"1"));
		currentAccount = pa.loadProperty(ACCOUNT_PROP_KEY, "");
		if (pa.loadProperty(FINISHED_FLAG_KEY, "TRUE") == "TRUE") {
			finished = true;
		}
		if (pa.loadProperty(FINISHED_FLAG_KEY, "TRUE") == "FALSE") {
			finished = false;
		}
	}

	private static void collectWeiboStatus(Timeline timeline) {
		Iterator<String> iScreenName = screenNameList.iterator();
		log.info("Starting Keyword Analysis...");
		// Fast-forward
		while (iScreenName.hasNext() && !currentScreenName.isEmpty()) {
			String screenName = iScreenName.next();
			if (screenName == currentScreenName) {
				break;
			}
		}
		File file = new File(OUTPUT_STATUS_FILE_NAME);
		try {
			FileOutputStream fos = new FileOutputStream(file);
			JsonWriter writer = new JsonWriter(new OutputStreamWriter(fos, "UTF-8"));
			writer.setIndent("    ");
			writer.beginArray();
			Gson gson = new Gson();
			
			while (iScreenName.hasNext()) {
				String screenName = null;
				
				if (finished || currentScreenName.isEmpty()) {
					screenName = iScreenName.next();
					currentScreenName = screenName;
				} else {
					screenName = currentScreenName;
				}
				log.info(String.format("Work on Screen Name[%s]", screenName));
				boolean hasNextPage = true;
				boolean outOfPeriod = false;
				int currentPageNum = 1;
				if (!finished) {
					currentPageNum = currentPageNumber;
				}
				StatusWapper statusWapper = getStatusWapper(timeline, screenName,
						currentPageNum);
				do {
					for (Status status : statusWapper.getStatuses()) {
	//					log.info(String.format("Screen Name[%s], Status[%s], Date[%s]", screenName, status.getText(), formatDate(status.getCreatedAt())));
						// Get the Last #Months Status
						Months months = calMthsBetween(status);
						if (months.getMonths() > PERIOD_OF_MONTHS) {
							outOfPeriod = true;
							// Out of period.
	//						log.info(String.format("Out of period. Status created at [%s], %s mths ago", formatDate(status.getCreatedAt()), months.getMonths()));
							break;
						}
//						keywordMatching(status.getText());
						Post post = new Post();
						post.setScreenName(screenName);
						post.setPostCreatedDate(status.getCreatedAt());
						post.setPostContent(status.getText());
						gson.toJson(post, Post.class, writer);
					}
	
					// Check if there is remaining posts
					hasNextPage = (statusWapper.getTotalNumber() - (currentPageNum * MAX_STATUS_CNT)) > 0;
					if (hasNextPage && !outOfPeriod) {
						currentPageNum++;
						statusWapper = getStatusWapper(timeline, screenName,
								currentPageNum);
						finished = false;
					} else {
						finished = true;
					}
					storeTempResult(screenName, currentPageNum);
				} while (hasNextPage && !outOfPeriod);
				
			}
			writer.endArray();
			writer.close();
			fos.close();
			log.info("Ended Keyword Analysis.");
		} catch (FileNotFoundException e) {
			log.error(e.getMessage(), e);
		} catch (UnsupportedEncodingException e) {
			log.error(e.getMessage(), e);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
	}
	
	private static void analysisWeiboStatuses(){
		log.info("Start to analysis keywords");
		Iterator<Post> iPost = statusesList.iterator();
		while(iPost.hasNext()){
			Post post = iPost.next();
			matchKeywords(post);
		}
	}
	
	private static void saveStatuses(){
		File file = new File(OUTPUT_STATUS_MODIFIED_FILE_NAME);
		log.info(String.format("Update Statuses Json File@%s", OUTPUT_STATUS_MODIFIED_FILE_NAME));
		try {
			FileOutputStream fos = new FileOutputStream(file);
			writerStatusesToJson(fos, statusesList);
			fos.close();
		} catch (FileNotFoundException e) {
			log.error(e.getMessage(), e);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
	}
	
	private static void writerStatusesToJson(OutputStream out, List<Post> statuses){
		try {
			JsonWriter writer = new JsonWriter(new OutputStreamWriter(out, "UTF-8"));
			Gson gson = new Gson();
			writer.setIndent("    ");
			writer.beginArray();
			for(Post post : statuses){
				gson.toJson(post, Post.class, writer);
			}
			writer.endArray();
			writer.close();
		} catch (UnsupportedEncodingException e) {
			log.error(e.getMessage(), e);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
	}
	
	private static void matchKeywords(Post post){
		Iterator<String> iKeyword = keywordList.iterator();
		while (iKeyword.hasNext()) {
			String keyword = iKeyword.next();
			Pattern pattern = Pattern.compile(Pattern
					.quote(keyword));
			Matcher matcher = pattern.matcher(post.getPostContent());
			if (matcher.find()) {
				post.getKeywords().add(keyword);
				List<Post> status = keywordStat.get(keyword);
				if(status == null){
					status = new ArrayList<Post>();
				}
				status.add(post);
				keywordStat.put(keyword, status);
			} 
		}
	}
	
	private static void stat(){
		Iterator<Post> iStatuses = statusesList.iterator();
		while(iStatuses.hasNext()){
			Post status = iStatuses.next();
			String dateKey = formatDate(status.getPostCreatedDate(), "yyyy/MM");
			Map<String, Map<String, Integer>> keywordsScreenNameMap = dateKeywordsNameMap.get(dateKey);
			if(keywordsScreenNameMap == null){
				keywordsScreenNameMap = new HashMap<String, Map<String, Integer>>();
			}
			List<String> keywords = status.getKeywords();
			if(keywords != null){
				Iterator<String> iKeywords = keywords.iterator();
				while(iKeywords.hasNext()){
					String keyword = iKeywords.next();
					Map<String, Integer> screenNameCntMap = keywordsScreenNameMap.get(keyword);
					if(screenNameCntMap == null){
						screenNameCntMap = new HashMap<String, Integer>();
					}
					Integer cnt = screenNameCntMap.get(status.getScreenName());
					screenNameCntMap.put(status.getScreenName(), (cnt == null) ? 1 : cnt+1);
					keywordsScreenNameMap.put(keyword, screenNameCntMap);
				}
			}
			dateKeywordsNameMap.put(dateKey, keywordsScreenNameMap);
		}
	}

	public static String formatDate(Date date) {
		return formatDate(date, "dd/MM/yyyy");
	}
	
	public static String formatDate(Date date, String format) {
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		return sdf.format(date);
	}

	private static void storeTempResult(String screenName, int currentPageNum) {
		storeResult(screenName, currentPageNum, currentAccount);
	}

//	private static void storeFinalResult() {
//		storeResult("", 0, "");
//	}

	private static void storeResult(String screenName, int currentPageNum,
			String currentAccount) {
		Map<String, String> propertiesMap = new HashMap<String, String>();
		propertiesMap.put(SCREEN_NAME_PROP_KEY, screenName);
		propertiesMap.put(PAGE_NUM_PROP_KEY, Integer.toString(currentPageNum));
		propertiesMap.put(ACCOUNT_PROP_KEY, currentAccount);
		propertiesMap.put(FINISHED_FLAG_KEY, finished ? "TRUE" : "FALSE");
		storeProperties(propertiesMap);
//		outputResult(nameListPath.getPath());
	}

	private static void storeProperties(Map<String, String> propertiesMap) {
		PropertiesAgent pa = new PropertiesAgent();
		pa.writeProperties(propertiesMap);
	}

	private static StatusWapper getStatusWapper(Timeline timeline,
			String screenName, int currentPageNum) {
		StatusWapper statusWapper = null;
		while(statusWapper == null){
			try {
				// Slow down the request frequency
				sleepFor(PAUSE_PERIOD_PER_REQUEST);
				statusWapper = timeline.getUserTimelineByName(screenName,
						new Paging(currentPageNum, MAX_STATUS_CNT), ALL_DATA,
						FEATURE_TYPE);
			} catch (WeiboException e) {
				if (e.getMessage().equals("api.weibo.com")){
					log.error(String.format("Weibo API server (%s) may be down OR unreachable", e.getMessage()), e);
				}
				// Out of limit for this account
				if (e.getErrorCode() == 10023) {
					// Change Token
					sleepForSwappingAccount();
					timeline.setToken(renewToken().getToken());
				} else {
					String msg = String.format("Error[%s],Account[%s]",
							e.getError(), currentAccount);
					log.error(msg, e);
				}
			} 
		}
		return statusWapper;
	}

	private static void sleepForSwappingAccount() {
		sleepFor(PAUSE_PERIOD);
	}

	private static void sleepFor(long millis) {
		try {
			// Pause
			log.info(String.format("Sleep for %d second...", millis / 1000));
			Thread.sleep(millis);
		} catch (InterruptedException e1) {
			e1.printStackTrace();
		}
	}

	private static void outputResult() {
		// Result
		log.info(String.format("Number of Status:%d", statusesList.size()));
		log.info(String.format("Number of Keywords:%d", keywordList.size()));
		log.info(String.format("Number of Account Monitored:%d", screenNameList.size()));
		log.info(String.format("Save file named:%s", OUTPUT_FILE_NAME));
		try {
			File outputfile = new File(OUTPUT_FILE_NAME);
			FileOutputStream fos = new FileOutputStream(outputfile);
			if (outputfile.exists()) {
				// Replace with a new file
				outputfile.delete();
			}
			WritableWorkbook outputWB = Workbook.createWorkbook(fos);
			
			// Simple Keyword counting
			WritableSheet sheet = outputWB.createSheet(OUTPUT_FILE_KEYWORD_SHEET_NAME, 0);
			Iterator<String> iKeyword = keywordList.iterator();

			int row = 1;
			sheet.addCell(new Label(OUTPUT_FILE_KEY_COL_IDX, 0,
					OUTPUT_FILE_KEY_COL_NAME));
			sheet.addCell(new Label(OUTPUT_FILE_VALUE_COL_IDX, 0,
					OUTPUT_FILE_VALUE_COL_NAME));
			sheet.addCell(new Label(OUTPUT_FILE_VALUE_COL_IDX+1, 0,
					OUTPUT_FILE_SCREEN_NAME_COL_NAME));
			while (iKeyword.hasNext()) {
				String keyword = iKeyword.next();
				Label keywordLb = new Label(OUTPUT_FILE_KEY_COL_IDX, row,
						keyword);
				Number freq = new Number(OUTPUT_FILE_VALUE_COL_IDX, row,
						(keywordStat.get(keyword) == null) ? 0
								: keywordStat.get(keyword).size());
				sheet.addCell(keywordLb);
				sheet.addCell(freq);
				Map<String, Integer> screenNameCnt = new HashMap<String, Integer>();
				if (keywordStat.get(keyword) != null && !keywordStat.get(keyword).isEmpty()){
					for(Post post : keywordStat.get(keyword)){
						Integer cnt = screenNameCnt.get(post.getScreenName());
						screenNameCnt.put(post.getScreenName(), (cnt == null) ? 1 : cnt+1);
					}
					int col = OUTPUT_FILE_VALUE_COL_IDX + 1;
					Iterator<String> iScreenNameCnt = screenNameCnt.keySet().iterator();
					while(iScreenNameCnt.hasNext()){
						String screenName = iScreenNameCnt.next();
						sheet.addCell(new Label(col++, row, String.format("%s[%d]", screenName, screenNameCnt.size())));
					}
				}
				row++;
			}
			
			// Date, Screen name keywords
			WritableSheet sheet2 = outputWB.createSheet(OUTPUT_FILE_DATE_KEYWORD_SHEET_NAME, 1);
			writeStatistic(sheet2, dateKeywordsNameMap,
					OUTPUT_FILE_DATE_COL_NAME,
					OUTPUT_FILE_KEY_COL_NAME,
					OUTPUT_FILE_MENTIONED_COL_NAME,
					OUTPUT_FILE_SCREEN_NAME_COL_NAME
					);
			outputWB.write();
			outputWB.close();
			fos.close();
		} catch (WriteException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void writeStatistic(WritableSheet sheet2, Map<String, Map<String, Map<String, Integer>>> stat,
			String col1Name, String col2Name, String col3Name, String col4Name)
			throws WriteException, RowsExceededException {
		int row;
		row = 1;
		sheet2.addCell(new Label(0, 0, col1Name));
		sheet2.addCell(new Label(1, 0, col2Name));
		sheet2.addCell(new Label(2, 0, col3Name));
		sheet2.addCell(new Label(3, 0, col4Name));
		Set<String> dates = stat.keySet();
		TreeSet<String> dateTree = new TreeSet<String>(dates);
		Iterator<String> iDSK = dateTree.iterator();
		while(iDSK.hasNext()){
			String dateKey = iDSK.next();
			sheet2.addCell(new Label(0, row, dateKey));
			Map<String, Map<String, Integer>> mapStrMap = stat.get(dateKey);
			if(mapStrMap == null){
				mapStrMap = new HashMap<String, Map<String, Integer>>();
			}
			Set<String> screenNames = mapStrMap.keySet();
			Iterator<String> iScreenNames = screenNames.iterator();
			while(iScreenNames.hasNext()){
				String screenNameKey = iScreenNames.next();
				sheet2.addCell(new Label(1, row, screenNameKey));
				Map<String, Integer> map = mapStrMap.get(screenNameKey);
				sheet2.addCell(new Number(2, row, map.size()));
				int col = 3;
				Iterator<String> iMap = map.keySet().iterator();
				while(iMap.hasNext()){
					String str = iMap.next();
					sheet2.addCell(new Label(col++, row, String.format("%s[%d]", str, map.get(str))));
				}
				row++;
			}
		}
	}

	private static URI getRes(String resFileName) {
		URL resPathURL = WeiboUpdated.class.getClassLoader().getResource(
				resFileName);
		URI resPath = null;
		if (resPathURL != null) {
			try {
				resPath = resPathURL.toURI();
			} catch (URISyntaxException e) {
				e.printStackTrace();
			}
		}
		return resPath;
	}

	private static Token renewToken() {
		Token accessToken = null;

		if (iAccounts.hasNext()) {
			currentAccount = iAccounts.next();
			log.info(String.format("Change to Account[%s]", currentAccount));
			String password = acctMap.get(currentAccount);
			try {
				accessToken = getToken(apiKey, apiSecret, currentAccount,
						password);
			} catch (FailingHttpStatusCodeException e) {
				e.printStackTrace();
			}
			if (!iAccounts.hasNext()) {
				// Renew account iterator
				iAccounts = accounts.iterator();
				log.info("Recycle Accounts!!!");
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
		Sheet sheet = null;
		try {
			keywordWB = Workbook.getWorkbook(file);
			sheet = keywordWB.getSheet(sheetIdx);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}

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
		Sheet sheet = null;
		try {
			screenNameWB = Workbook.getWorkbook(file);
			sheet = screenNameWB.getSheet(sheetIdx);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}

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
		String authorizationUrl = service.getAuthorizationUrl(EMPTY_TOKEN);
		log.debug(String.format("authorizationUrl[%s]", authorizationUrl));

		String access_code = null;
		try {
			access_code = getAccessCode(email, password, authorizationUrl);
		} catch (FailingHttpStatusCodeException e) {
			log.error("Fail to get access code", e);
		}
		log.debug(String.format("access_code[%s]", access_code));
		Verifier verifier = new Verifier(access_code);

		return service.getAccessToken(EMPTY_TOKEN, verifier);
	}

	public static String getAccessCode(final String email,
			final String password, final String authorizationUrl)
			throws FailingHttpStatusCodeException {
		log.debug(String
				.format("Start to get access code with Account[%s],AuthorizationURL[%s]",
						email, authorizationUrl));
		String accessCode = null;

		FirefoxDriver driver = new FirefoxDriver();

		driver.get(authorizationUrl);

		WebElement userIdInput = driver.findElement(By.name("userId"));
		WebElement passwdInput = driver.findElement(By.name("passwd"));
		WebElement vcodeSectoin = driver
				.findElementByXPath(VCODE_SECTION_XPATH);

		userIdInput.sendKeys(email, Keys.TAB);
		passwdInput.sendKeys(password, Keys.TAB);
		passwdInput.submit();
		try{
			if (vcodeSectoin.getAttribute("style").isEmpty()) {
				WebDriverWait waiter = new WebDriverWait(driver, TIME_OUT_SECONDS, WAIT_FOR_INPUT_MILLIS);
				waiter.until(ExpectedConditions.titleIs(PAGE_TITLE));
			}
		}catch(StaleElementReferenceException e){
			log.error(String.format("Can Auto-login, since [%s]", e.getMessage()), e);
		}

		String urlStr = driver.getCurrentUrl();

		driver.close();

		accessCode = urlStr.substring(urlStr.length() - ACCESS_CODE_LENGTH);

		if (accessCode != null && accessCode.length() > 32) {
			log.debug(String.format("accessCode[%s]", accessCode));
			sleepForSwappingAccount();
			return getAccessCode(email, password, authorizationUrl);
		}
		return accessCode;
	}

}
