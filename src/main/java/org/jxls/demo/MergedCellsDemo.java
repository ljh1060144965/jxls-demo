package org.jxls.demo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.demo.util.JxlsUtils;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Leonid Vysochyn on 6/30/2015.
 * todo: improve each command to be able to set merge cells
 */
public class MergedCellsDemo  {
    static Logger logger = LoggerFactory.getLogger(MergedCellsDemo.class);
    private static String template = "merged_cells_demo.xls";
    private static String output = "target/merged_cells_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Running merged cells demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = XlsCommentBuilderDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Map<String, Object> context = new HashMap<>();
                context.put("departments", departments);
//
                Workbook wb = JxlsUtils.buildExcel(is, context);
                Sheet sheet = wb.getSheetAt(0);
                Row row = sheet.getRow(0);
                row.getCell(0).setCellValue("xxxxxxxxxx2");
                wb.write(new FileOutputStream(output));

            }
        }
    }
}
