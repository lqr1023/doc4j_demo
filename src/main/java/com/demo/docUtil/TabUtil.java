package com.demo.docUtil;

import java.math.BigInteger;
import java.util.List;

import org.docx4j.jaxb.Context;
import org.docx4j.model.properties.table.tr.TrHeight;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTHeight;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.CTVerticalJc;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STHint;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.TrPr;
import org.junit.Test;

public class TabUtil {
	public void createTab() throws Exception {
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		int writableWidthTwips = wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions()
				.getWritableWidthTwips();
		int columnNumber = 3;
		Tbl tbl = TblFactory.createTable(3, 3, writableWidthTwips / columnNumber);
		List<Object> rows = tbl.getContent();
		for (Object row : rows) {
			Tr tr = (Tr) row;
			List<Object> cells = tr.getContent();
			for (Object cell : cells) {
				Tc td = (Tc) cell;
				
				//td.getContent().add();
			}
		}
	}
	
	public void createCell(){
		
	}
	
	// 表格增加边框  
    public void addBorders(Tbl table,CTBorder border){  
        table.setTblPr(new TblPr());  
        TblBorders borders = new TblBorders();  
        borders.setBottom(border);  
        borders.setLeft(border);  
        borders.setRight(border);  
        borders.setTop(border);  
        borders.setInsideH(border);  
        borders.setInsideV(border);  
        table.getTblPr().setTblBorders(borders);  
    }  
  
    // 表格增加边框 可以设置上下左右四个边框样式以及横竖水平线样式  
    public void addBorders(Tbl table, CTBorder topBorder,  
            CTBorder bottomBorder, CTBorder leftBorder, CTBorder rightBorder,  
            CTBorder hBorder, CTBorder vBorder) {  
        table.setTblPr(new TblPr());  
        TblBorders borders = new TblBorders();  
        borders.setBottom(bottomBorder);  
        borders.setLeft(leftBorder);  
        borders.setRight(rightBorder);  
        borders.setTop(bottomBorder);  
        borders.setInsideH(hBorder);  
        borders.setInsideV(vBorder);  
        table.getTblPr().setTblBorders(borders);  
    }  
	
    // 得到页面宽度  
    public int getWritableWidth(WordprocessingMLPackage wordPackage)  
            throws Exception {  
        return wordPackage.getDocumentModel().getSections().get(0)  
                .getPageDimensions().getWritableWidthTwips();  
    }
    
    /** 
     * @param tableWidthPercent 
     *            表格占页面宽度百分比 
     * @param widthPercent 
     *            各列百分比 
     */  
    public void setTableGridCol(WordprocessingMLPackage wordPackage,  
            ObjectFactory factory, Tbl table, double tableWidthPercent,  
            double[] widthPercent) throws Exception {  
        int width = getWritableWidth(wordPackage);  
        int tableWidth = (int) (width * tableWidthPercent / 100);  
        TblGrid tblGrid = factory.createTblGrid();  
        for (int i = 0; i < widthPercent.length; i++) {  
            TblGridCol gridCol = factory.createTblGridCol();  
            gridCol.setW(BigInteger.valueOf((long) (tableWidth  
                    * widthPercent[i] / 100)));  
            tblGrid.getGridCol().add(gridCol);  
        }  
        table.setTblGrid(tblGrid);  
  
        TblPr tblPr = table.getTblPr();  
        if (tblPr == null) {  
            tblPr = factory.createTblPr();  
        }  
        TblWidth tblWidth = new TblWidth();  
        tblWidth.setType("dxa");// 这一行是必须的,不自己设置宽度默认是auto  
        tblWidth.setW(new BigInteger(tableWidth + ""));  
        tblPr.setTblW(tblWidth);  
        table.setTblPr(tblPr);  
    }  
    
    // 设置tr高度  
    public void setTableTrHeight(ObjectFactory factory, Tr tr, String heigth) {  
        TrPr trPr = tr.getTrPr();  
        if (trPr == null) {  
            trPr = factory.createTrPr();  
        }  
        CTHeight ctHeight = new CTHeight();  
        ctHeight.setVal(new BigInteger(heigth));  
        TrHeight trHeight = new TrHeight(ctHeight);  
        trHeight.set(trPr);  
        tr.setTrPr(trPr);  
    }
    
    // 新增单元格  
    public void addTableCell(ObjectFactory factory,  
            WordprocessingMLPackage wordMLPackage, Tr tableRow, String content,  
            RPr rpr, JcEnumeration jcEnumeration, boolean hasBgColor,  
            String backgroudColor) {  
        Tc tableCell = factory.createTc();  
        P p = factory.createP();  
        ParUtil.setParagraphSpacing(factory, p, jcEnumeration,  "0", "0", null,  
                null, "240", STLineSpacingRule.AUTO);  
        Text t = factory.createText();  
        t.setValue(content);  
        R run = factory.createR();  
        // 设置表格内容字体样式  
        run.setRPr(rpr);  
  
        TcPr tcPr = tableCell.getTcPr();  
        if (tcPr == null) {  
            tcPr = factory.createTcPr();  
        }  
  
        CTVerticalJc valign = factory.createCTVerticalJc();  
        valign.setVal(STVerticalJc.CENTER);  
        tcPr.setVAlign(valign);  
  
        run.getContent().add(t);  
        p.getContent().add(run);  
        tableCell.getContent().add(p);  
        if (hasBgColor) {  
            CTShd shd = tcPr.getShd();  
            if (shd == null) {  
                shd = factory.createCTShd();  
            }  
            shd.setColor("auto");  
            shd.setFill(backgroudColor);  
            tcPr.setShd(shd);  
        }  
        tableCell.setTcPr(tcPr);  
        tableRow.getContent().add(tableCell);  
    }
    
    // 表格水平对齐方式  
    public void setTableAlign(ObjectFactory factory, Tbl table,  
            JcEnumeration jcEnumeration) {  
        TblPr tablePr = table.getTblPr();  
        if (tablePr == null) {  
            tablePr = factory.createTblPr();  
        }  
        Jc jc = tablePr.getJc();  
        if (jc == null) {  
            jc = new Jc();  
        }  
        jc.setVal(jcEnumeration);  
        tablePr.setJc(jc);  
        table.setTblPr(tablePr);  
    }  
    
    @Test
	public void createNoMergeTab() throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
		ObjectFactory factory = Context.getWmlObjectFactory();
		RPr titleFont = ParUtil.getRPr(factory, "宋体", "000000", "22", STHint.EAST_ASIA, true, false, false, false);
		RPr contentFont = ParUtil.getRPr(factory, "宋体", "000000", "22", STHint.EAST_ASIA, false, false, false, false);
		//表格主体
		Tbl table = factory.createTbl();
		//添加边框
		CTBorder border = new CTBorder(); 
		border.setColor("red");
		border.setVal(STBorder.SINGLE);
		border.setSz(new BigInteger("2"));
		addBorders(table,border);
		double[] colWidthPercent = new double[] { 15, 20, 20, 20};// 百分比  
		setTableGridCol(wordMLPackage, factory, table, 80, colWidthPercent); 
		Tr titleRow = factory.createTr();  
        setTableTrHeight(factory, titleRow, "500");
        addTableCell(factory, wordMLPackage, titleRow, "序号", titleFont,  
                JcEnumeration.CENTER, true, "C6D9F1");  
        addTableCell(factory, wordMLPackage, titleRow, "姓甚", titleFont,  
                JcEnumeration.CENTER, true, "C6D9F1");  
        addTableCell(factory, wordMLPackage, titleRow, "名谁", titleFont,  
                JcEnumeration.CENTER, true, "C6D9F1");  
        addTableCell(factory, wordMLPackage, titleRow, "籍贯", titleFont,  
                JcEnumeration.CENTER, true, "C6D9F1");  
        addTableCell(factory, wordMLPackage, titleRow, "营生", titleFont,  
                JcEnumeration.CENTER, true, "C6D9F1");  
        table.getContent().add(titleRow);  
        for (int i = 0; i < 10; i++) {  
            Tr contentRow = factory.createTr();
            setTableTrHeight(factory, contentRow, "500");
            addTableCell(factory, wordMLPackage, contentRow, i + "",  
                    contentFont, JcEnumeration.CENTER, false, null);  
            addTableCell(factory, wordMLPackage, contentRow, "无名氏", contentFont,  
                    JcEnumeration.CENTER, false, null);  
            addTableCell(factory, wordMLPackage, contentRow, "佚名", contentFont,  
                    JcEnumeration.CENTER, false, null);  
            addTableCell(factory, wordMLPackage, contentRow, "武林", contentFont,  
                    JcEnumeration.CENTER, false, null);  
            addTableCell(factory, wordMLPackage, contentRow, "吟诗赋曲",  
            		contentFont, JcEnumeration.CENTER, false, null);  
            table.getContent().add(contentRow);  
        }  
        setTableAlign(factory, table, JcEnumeration.LEFT);  
        
        mp.addObject(table);
        
        wordMLPackage.save(new java.io.File("d:/HelloWord1.docx"));
        
	}
}
