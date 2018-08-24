package com.demo.docUtil;

import java.math.BigInteger;


import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STHint;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.Text;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;

import com.demo.entity.Font;
import com.demo.entity.ParStyle;

import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.Spacing;


public class ParUtil {
	
	/** 
     * 创建字体 
     *  
     * @param isBlod 
     *            粗体 
     * @param isUnderLine 
     *            下划线 
     * @param isItalic 
     *            斜体 
     * @param isStrike 
     *            删除线 
     */  
    public static RPr getRPr(ObjectFactory factory, String fontFamily,  
            String colorVal, String fontSize, STHint sTHint, boolean isBlod,  
            boolean isUnderLine, boolean isItalic, boolean isStrike) {  
        RPr rPr = factory.createRPr();  
        RFonts rf = new RFonts();  
        rf.setHint(sTHint);  
        rf.setAscii(fontFamily);  
        rf.setHAnsi(fontFamily);  
        rPr.setRFonts(rf);  
  
        BooleanDefaultTrue bdt = factory.createBooleanDefaultTrue();  
        rPr.setBCs(bdt);  
        if (isBlod) {  
            rPr.setB(bdt);  
        }  
        if (isItalic) {  
            rPr.setI(bdt);  
        }  
        if (isStrike) {  
            rPr.setStrike(bdt);  
        }  
        if (isUnderLine) {  
            U underline = new U();  
            underline.setVal(UnderlineEnumeration.SINGLE);  
            rPr.setU(underline);  
        }  
  
        Color color = new Color();  
        color.setVal(colorVal);  
        rPr.setColor(color);  
  
        HpsMeasure sz = new HpsMeasure();  
        sz.setVal(new BigInteger(fontSize));  
        rPr.setSz(sz);  
        rPr.setSzCs(sz);  
        return rPr;  
    }  
	

 // 设置段间距-->行距 段前段后距离  
    // 段前段后可以设置行和磅 行距只有磅  
    // 段前磅值和行值同时设置，只有行值起作用  
    // TODO 1磅=20 1行=100 单倍行距=240 为什么是这个值不知道  
    /** 
     * @param jcEnumeration 
     *            对齐方式 
     * @param isSpace 
     *            是否设置段前段后值 
     * @param before 
     *            段前磅数 
     * @param after 
     *            段后磅数 
     * @param beforeLines 
     *            段前行数 
     * @param afterLines 
     *            段后行数 
     * @param isLine 
     *            是否设置行距 
     * @param lineValue 
     *            行距值 
     * @param sTLineSpacingRule 
     *            自动auto 固定exact 最小 atLeast 
     */  
    public static void setParagraphSpacing(ObjectFactory factory, P p,
    		JcEnumeration jcEnumeration,String before,String after,
    		String beforeLines, String afterLines,String lineValue,  
            STLineSpacingRule sTLineSpacingRule) {  
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

		Spacing spacing = new Spacing();
		if (before != null) {
			// 段前磅数
			spacing.setBefore(new BigInteger(before));
		}
		if (after != null) {
			// 段后磅数
			spacing.setAfter(new BigInteger(after));
		}
		if (beforeLines != null) {
			// 段前行数
			spacing.setBeforeLines(new BigInteger(beforeLines));
		}
		if (afterLines != null) {
			// 段后行数
			spacing.setAfterLines(new BigInteger(afterLines));
		}

		if (lineValue != null) {
			spacing.setLine(new BigInteger(lineValue));
		}
		spacing.setLineRule(sTLineSpacingRule);

		pPr.setSpacing(spacing);
		p.setPPr(pPr);
    }  
  
    // 设置缩进 同时设置为true,则为悬挂缩进  
    public static void setParagraphInd(ObjectFactory factory, P p,String firstLineValue, String hangValue) {  
        PPr pPr = p.getPPr();  
        if (pPr == null) {  
            pPr = factory.createPPr();  
        }  
        Jc jc = pPr.getJc();  
        if (jc == null) {  
            jc = new Jc();  
        }  
        pPr.setJc(jc);  
  
        Ind ind = pPr.getInd();  
        if (ind == null) {  
            ind = new Ind();  
        }  
        if (firstLineValue != null) {  
                ind.setFirstLineChars(new BigInteger(firstLineValue));  
        }  
         
        if (hangValue != null) {  
                ind.setHangingChars(new BigInteger(hangValue));  
        }  
        pPr.setInd(ind);  
        p.setPPr(pPr);  
    }  				
    
	
	/**
	 * 
	 * @param factory
	 * @param mp
	 * @param font 字体实体类
	 * @param ps 段落格式
	 * @param text 填充文字
	 */
	public static void createPar(ObjectFactory factory,MainDocumentPart mp,Font font,ParStyle ps,String text){
    	RPr rPr = getRPr(factory, font.getFontFamily(), font.getColorVal(), font.getFontSize(), font.getsTHint(), font.isBlod(), font.isUnderLine(), font.isItalic(), font.isStrike());
    	P paragraph = factory.createP();
    	Text txt = factory.createText();
    	R run = factory.createR();
    	txt.setValue(text);
    	run.getContent().add(txt);
		run.setRPr(rPr);
		paragraph.getContent().add(run);
		setParagraphSpacing(factory, paragraph, ps.getJcEnumeration(), ps.getBefore(), ps.getAfter(), ps.getBeforeLines(), ps.getAfterLines(), ps.getLineValue(), ps.getsTLineSpacingRule());
		setParagraphInd(factory,paragraph,ps.getLineValue(),ps.getHangValue());
		mp.addObject(paragraph);
    }

}
