package org.jxls.demo.guide;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.common.Context;
import org.jxls.demo.util.JxlsUtils;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Object collection output demo
 * @author Leonid Vysochyn
 */
public class ObjectCollectionDemo {
    private static Logger logger = LoggerFactory.getLogger(ObjectCollectionDemo.class);

    public static void main(String[] args) throws ParseException, IOException, InvalidFormatException {
        logger.info("Running Object Collection demo");
        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {
                Map<String, Object> context = new HashMap<>();
                context.put("employees", employees);
                Workbook workbook = JxlsUtils.buildExcel(is, context);
                Sheet sheet = workbook.getSheetAt(0);
                //第一个参数为合并起始行，从0开始
                //第二个参数为合并终止行，从0开始
                //第二个参数为合并起始列，从0开始
                //第二个参数为合并终止列，从0开始
                //例子中是合并第一行的1列---8列
//                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 7);
//                CellRangeAddress region = new CellRangeAddress(3, 4, 0, 0);
//                sheet.addMergedRegion(region);
//                sheet.getRow(5).setHeightInPoints(32.5f);
                workbook.write(os);
            }
        }
    }

    public static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("李佳辉", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("李佳辉", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        return employees;
    }
}
