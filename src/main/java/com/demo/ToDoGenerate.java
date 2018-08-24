package com.demo;

import org.docx4j.openpackaging.exceptions.Docx4JException;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;

/**
 * 
 * 
 * word文档解析的是xml文件，docx4j可以通过操作xml来完成文档的制作
 * 测试docx4j生成word文档
 * 解析xml/json
 * 段落 含有复杂格式
 * 加图片
 * 表格
 * 排序添加参考文献
 * 
 * @author User
 *
 */
public class ToDoGenerate {
	@Test
	public void simple() throws Exception{
		//最简单的实例代码 在word文档里添加文字
		//获取doc xml文件包
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		//获取主体部分
		wordMLPackage.getMainDocumentPart().addParagraphOfText("Hello Word!");
		//保存文件
		wordMLPackage.save(new java.io.File("d:/HelloWord1.docx"));
	}
	
	@Test
	public void docwithStyle() throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		//增加样式文件
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "Hello World");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Subtitle", "This is a subTitle");
		//保存文件
		wordMLPackage.save(new java.io.File("d:/HelloWord1.docx"));
		
	}

}
