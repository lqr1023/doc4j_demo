package com.demo;

import java.math.BigInteger;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;

public class TableWithBorders {
    private static WordprocessingMLPackage  wordMLPackage;
    private static ObjectFactory factory;
 
    public static void main (String[] args) throws Docx4JException {
        wordMLPackage = WordprocessingMLPackage.createPackage();
        factory = Context.getWmlObjectFactory();
 
        Tbl table = createTableWithContent();
 
        addBorders(table);
 
        wordMLPackage.getMainDocumentPart().addObject(table);
        wordMLPackage.save(new java.io.File(
            "d:/HelloWord3.docx") );
    }
 
    private static void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("red");
        border.setSz(new BigInteger("10"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);
 
        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }
 
    private static Tbl createTableWithContent() {
        Tbl table = factory.createTbl();
        Tr tableRow = factory.createTr();
        addTableCell(tableRow, "Field 1");
        addTableCell(tableRow, "Field 2");
        table.getContent().add(tableRow);
        return table;
    }
 
    private static void addTableCell(Tr tableRow, String content) {
        Tc tableCell = factory.createTc();
        tableCell.getContent().add(
        wordMLPackage.getMainDocumentPart().
            createParagraphOfText(content));
        tableRow.getContent().add(tableCell);
    }
}
