package com.demo.entity;


import org.docx4j.wml.STHint;

/**
 * 创建doc文件   字体实体类
 * @author User
 *
 */
public class Font {
	/**
	 * 字体
	 */
	private String fontFamily;
	/**
	 * RGB 六位颜色数值
	 * https://www.rapidtables.com/web/color/RGB_Color.html
	 */
	private String colorVal;
	/**
	 * 字号 doc文档中字号的2倍
	 */
	private String fontSize;
	
	/**
	 * 字体样式 
	 *  STHint.CS
	 *  STHint.DEFAULT
	 * 	STHint.EAST_ASIA
	 */
	private STHint sTHint;
	/**
	 * 是否加粗
	 */
	private boolean isBlod;
	/**
	 * 是否加下划线
	 */
	private boolean isUnderLine;
	/**
	 * 是否斜体
	 */
	private boolean isItalic;
	/**
	 * 是否加删除线 
	 */
	private boolean isStrike;
	
	
	public Font(String fontFamily, String colorVal, String fontSize, STHint sTHint, boolean isBlod, boolean isUnderLine,
			boolean isItalic, boolean isStrike) {
		super();
		this.fontFamily = fontFamily;
		this.colorVal = colorVal;
		this.fontSize = fontSize;
		this.sTHint = sTHint;
		this.isBlod = isBlod;
		this.isUnderLine = isUnderLine;
		this.isItalic = isItalic;
		this.isStrike = isStrike;
	}
	public String getFontFamily() {
		return fontFamily;
	}
	public void setFontFamily(String fontFamily) {
		this.fontFamily = fontFamily;
	}
	public String getColorVal() {
		return colorVal;
	}
	public void setColorVal(String colorVal) {
		this.colorVal = colorVal;
	}
	public String getFontSize() {
		return fontSize;
	}
	public void setFontSize(String fontSize) {
		this.fontSize = fontSize;
	}
	public boolean isBlod() {
		return isBlod;
	}
	public void setBlod(boolean isBlod) {
		this.isBlod = isBlod;
	}
	public boolean isUnderLine() {
		return isUnderLine;
	}
	public void setUnderLine(boolean isUnderLine) {
		this.isUnderLine = isUnderLine;
	}
	public boolean isItalic() {
		return isItalic;
	}
	public void setItalic(boolean isItalic) {
		this.isItalic = isItalic;
	}
	public boolean isStrike() {
		return isStrike;
	}
	public void setStrike(boolean isStrike) {
		this.isStrike = isStrike;
	}
	public STHint getsTHint() {
		return sTHint;
	}
	public void setsTHint(STHint sTHint) {
		this.sTHint = sTHint;
	}
	
	
}
