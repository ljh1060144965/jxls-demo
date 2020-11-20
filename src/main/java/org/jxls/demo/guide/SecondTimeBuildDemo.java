package org.jxls.demo.guide;

import jxl.CellReferenceHelper;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.jxls.demo.util.JxlsUtils;
import org.jxls.util.CellRefUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
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
 * 第一次build玩表格后，会将表格的注释comment清除掉，就会导致二次读取build时会导致找不到area的注释comment，无法进行二次build，
 * 所以二次读取时需手动编码设置area域和相关jxls表达式
 *
 * @author Leonid Vysochyn
 */
public class SecondTimeBuildDemo {
    private static Logger logger = LoggerFactory.getLogger(SecondTimeBuildDemo.class);

    public static void main(String[] args) throws ParseException, IOException, InvalidFormatException {
        logger.info("Running Object Collection demo");
        List<Employee> employees = generateSampleEmployeeData();
        try(
                InputStream is = SecondTimeBuildDemo.class.getResourceAsStream("secondtime_build_template.xls")
        ) {
            try (OutputStream os = new FileOutputStream("target/secondtime_build_template33333.xls")) {
                Map<String, Object> context = new HashMap<>();
                context.put("time1", "2020-05-02多次build112");
                Workbook workbook = JxlsUtils.buildExcel(is, context);
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
