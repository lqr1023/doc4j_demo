package com.demo.docUtil;

import java.io.File;
import java.nio.file.Files;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.Text;

public class PicUtil {
	// 段落中插入文字和图片  
    public static P createImageParagraph(WordprocessingMLPackage wordMLPackage,  
            ObjectFactory factory, P p, String fileName, String content,  
            byte[] bytes, JcEnumeration jcEnumeration) throws Exception {  
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage  
                .createImagePart(wordMLPackage, bytes);  
        Inline inline = imagePart.createImageInline(fileName, "这是图片", 1, 2,  
                false);  
        Text text = factory.createText();  
        text.setValue(content);  
        text.setSpace("preserve");  
        R run = factory.createR();  
        p.getContent().add(run);  
        run.getContent().add(text);  
        Drawing drawing = factory.createDrawing();  
        run.getContent().add(drawing);  
        drawing.getAnchorOrInline().add(inline);  
        PPr pPr = p.getPPr();  
        if (pPr == null) {  
            pPr = factory.createPPr();  
        }  
        Jc jc = pPr.getJc();  
        if (jc == null) {  
            jc = new Jc();  
        }  
        jc.setVal(jcEnumeration);  
        pPr.setJc(jc);  
        p.setPPr(pPr);  
        ParUtil.setParagraphSpacing(factory, p, jcEnumeration,"0", "0", null, null, "240", STLineSpacingRule.AUTO);  
        return p;  
    }  
	
    public static void addImageToParagraph(WordprocessingMLPackage wordMLPackage,String fileName) throws Exception {
    	MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
    	File image = new File(fileName);
    	byte[] fileContent = Files.readAllBytes(image.toPath());
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage  
                .createImagePart(wordMLPackage, fileContent);  
        Inline inline = imagePart.createImageInline(fileName, "这是图片", 1, 2,  
                false);
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        mp.getContent().add(p);
    }


}
