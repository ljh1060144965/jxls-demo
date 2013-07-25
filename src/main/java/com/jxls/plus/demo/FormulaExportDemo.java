package com.jxls.plus.demo;

import com.jxls.plus.area.XlsArea;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12
 */
public class FormulaExportDemo {
    static Logger logger = LoggerFactory.getLogger(FormulaExportDemo.class);
    private static String template = "formulas_demo.xlsx";
    private static String output = "target/formulas_demo_output.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing formulas demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        logger.info("Opening input stream");
        InputStream is = FormulaExportDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = PoiTransformer.createTransformer(is, os);
        XlsArea sheet1Area = new XlsArea("Sheet1!A1:D4", transformer);
        XlsArea sheet2Area = new XlsArea("Sheet2!A1:A2", transformer);
        XlsArea sheet3Area = new XlsArea("'Sheet 3'!A1:A2", transformer);
        Context context = new Context();
        sheet3Area.applyAt(new CellRef("Sheet1!K1"), context);
        sheet2Area.applyAt(new CellRef("Sheet2!B6"), context);
        sheet2Area.applyAt(new CellRef("Sheet2!C6"), context);
        sheet2Area.applyAt(new CellRef("Sheet2!D6"), context);
        sheet1Area.applyAt(new CellRef("Sheet1!F11"), context);
        sheet1Area.processFormulas();
        sheet1Area.clearCells();
        transformer.write();
        logger.info("written to file");
        is.close();
        os.close();
    }

}
