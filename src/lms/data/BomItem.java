/**
 * 
 */
package lms.data;

/**
 * @author xusu1
 *
 */
public class BomItem {
	private 	String Number;
	private String BomCode;
	private String FeatureCode;
	private String FeatureCodeDesc;
	private String amount;
	
	public BomItem(String nu,String bomCode,String featureCode,String fearureCodeDesc){
		this.Number=nu;
		this.BomCode=bomCode;
		this.FeatureCode=featureCode;
		this.FeatureCodeDesc=fearureCodeDesc;
	}
	
	
	public BomItem() {
		// TODO Auto-generated constructor stub
	}


	/**
	 * @return the number
	 */
	
	
	public String getNumber() {
		return Number;
	}
	/**
	 * @param number the number to set
	 */
	public void setNumber(String number) {
		Number = number;
	}
	/**
	 * @return the bomCode
	 */
	public String getBomCode() {
		return BomCode;
	}
	/**
	 * @param bomCode the bomCode to set
	 */
	public void setBomCode(String bomCode) {
		BomCode = bomCode;
	}
	/**
	 * @return the featureCode
	 */
	public String getFeatureCode() {
		return FeatureCode;
	}
	/**
	 * @param featureCode the featureCode to set
	 */
	public void setFeatureCode(String featureCode) {
		FeatureCode = featureCode;
	}
	/**
	 * @return the featureCodeDesc
	 */
	public String getFeatureCodeDesc() {
		return FeatureCodeDesc;
	}
	/**
	 * @param featureCodeDesc the featureCodeDesc to set
	 */
	public void setFeatureCodeDesc(String featureCodeDesc) {
		FeatureCodeDesc = featureCodeDesc;
	}
	/**
	 * @return the amount
	 */
	public String getAmount() {
		return amount;
	}
	/**
	 * @param amount the amount to set
	 */
	public void setAmount(String amount) {
		this.amount = amount;
	}


	@Override
	public String toString() {
		return "BomItem [Number=" + Number + ", BomCode=" + BomCode
				+ ", FeatureCode=" + FeatureCode + ", FeatureCodeDesc="
				+ FeatureCodeDesc + ", amount=" + amount + "]\n";
	}

}
