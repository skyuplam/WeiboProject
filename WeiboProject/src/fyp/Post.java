package fyp;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Post implements java.io.Serializable{
	private static final long serialVersionUID = 3504482275852201185L;
	private String screenName;
	private String postContent;
	private Date postCreatedDate;
	private List<String> keywords = new ArrayList<String>();
	
	public String getScreenName() {
		return screenName;
	}
	public void setScreenName(String screenName) {
		this.screenName = screenName;
	}
	public String getPostContent() {
		return postContent;
	}
	public void setPostContent(String postContent) {
		this.postContent = postContent;
	}
	public Date getPostCreatedDate() {
		return postCreatedDate;
	}
	public void setPostCreatedDate(Date postCreatedDate) {
		this.postCreatedDate = postCreatedDate;
	}
	public List<String> getKeywords() {
		return keywords;
	}
	public void setKeywords(List<String> keywords) {
		this.keywords = keywords;
	}
}
