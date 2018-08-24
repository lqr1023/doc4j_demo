package com.demo.mainPro;

import java.io.File;


import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.STHint;
import org.docx4j.wml.STLineSpacingRule;

import com.demo.docUtil.ParUtil;
import com.demo.docUtil.PicUtil;
import com.demo.entity.Font;
import com.demo.entity.ParStyle;

/**
 * 参考资料：
 * http://53873039oycg.iteye.com/blog/2123815
 * https://www.baeldung.com/docx4j
 * @author User
 *
 */
public class CreateDoc {
	
	public static void main(String[] args) throws Exception {
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
		ObjectFactory factory = Context.getWmlObjectFactory();
		Font font = new Font("宋体", "000000", "18", STHint.EAST_ASIA, false, false, false, false);
		ParStyle ps = new ParStyle(JcEnumeration.LEFT,null,"0", null, null, "240", STLineSpacingRule.AUTO, "20", null);
		ParUtil.createPar(factory, mp, font,ps, "Docx4j中,在内存中操作的word文档是\"WordprocessingMLPackage\"类型对象(本文一下简称包) 在编辑一个word文档前,开发者需要选择:创建一个新的空白包,并逐一将需要的内容填充进去,或者打开一个已有的文档,并在里面添加/替换新的内容前者思路比较简单,比较适合简单文档的创建.但由于添加每条新内容时,都需要手动进行设置其各项参数(比如表格的行款,列宽,边框等),且添加修改复杂空间(公式,页眉页脚)的过程都比较繁琐,所以在创建格式复杂的文档时不是很建议.后者需要事先制作一个模板文档,添加不同的占位符和各种模板信息,在准备上比前者复杂,但也具有很多优点:可以简化细节参数的调整(不需要手动调整表格,段落的具体细节参数)从而将精力集中到文档内容上复杂的文档部分(如,公式/复选框等)可以直接从模板中读取,只需要在其基础上修改文字等内容部分,而避开了繁琐的创建操作等在创建格式复杂的文档时,这个方法相比前者可以精简大量代码");
		ParUtil.createPar(factory, mp, font,ps,"打开文件,通过imagePart将图片读进去,现在图片被转换成二进制,为了能在文件内联中显示出图片,调用函数将图片存放在inline中,之后paragraph,run中可以用drawing读取inline");
		PicUtil.addImageToParagraph(wordMLPackage, "src/main/java/com/demo/imgs/1.png");
		wordMLPackage.save(new java.io.File("d:/HelloWord1.docx"));
	}

}
