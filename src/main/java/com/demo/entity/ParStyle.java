package com.demo.entity;

import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.STLineSpacingRule;

public class ParStyle {
	
	//JcEnumeration.LEFT
	private JcEnumeration jcEnumeration;
	//段前距离
	private String before;
	//段后距离
	private String after;
	//段前行数 
	private String beforeLines;
	//段后行数
	private String afterLines;
	//行距
	private String lineValue;
	//自动auto 固定exact 最小 atLeast
	private STLineSpacingRule sTLineSpacingRule;
	//首行缩进
	private String firstLineValue;
	//悬挂缩进
	private String hangValue;

	public ParStyle(JcEnumeration jcEnumeration, String before, String after, String beforeLines, String afterLines,
			String lineValue, STLineSpacingRule sTLineSpacingRule, String firstLineValue, String hangValue) {
		super();
		this.jcEnumeration = jcEnumeration;
		this.before = before;
		this.after = after;
		this.beforeLines = beforeLines;
		this.afterLines = afterLines;
		this.lineValue = lineValue;
		this.sTLineSpacingRule = sTLineSpacingRule;
		this.firstLineValue = firstLineValue;
		this.hangValue = hangValue;
	}


	public JcEnumeration getJcEnumeration() {
		return jcEnumeration;
	}


	public void setJcEnumeration(JcEnumeration jcEnumeration) {
		this.jcEnumeration = jcEnumeration;
	}


	public String getBefore() {
		return before;
	}


	public void setBefore(String before) {
		this.before = before;
	}


	public String getAfter() {
		return after;
	}


	public void setAfter(String after) {
		this.after = after;
	}


	public String getBeforeLines() {
		return beforeLines;
	}


	public void setBeforeLines(String beforeLines) {
		this.beforeLines = beforeLines;
	}


	public String getAfterLines() {
		return afterLines;
	}


	public void setAfterLines(String afterLines) {
		this.afterLines = afterLines;
	}


	public String getLineValue() {
		return lineValue;
	}


	public void setLineValue(String lineValue) {
		this.lineValue = lineValue;
	}


	public STLineSpacingRule getsTLineSpacingRule() {
		return sTLineSpacingRule;
	}


	public void setsTLineSpacingRule(STLineSpacingRule sTLineSpacingRule) {
		this.sTLineSpacingRule = sTLineSpacingRule;
	}


	public String getFirstLineValue() {
		return firstLineValue;
	}


	public void setFirstLineValue(String firstLineValue) {
		this.firstLineValue = firstLineValue;
	}


	public String getHangValue() {
		return hangValue;
	}

	public void setHangValue(String hangValue) {
		this.hangValue = hangValue;
	}
	
}
