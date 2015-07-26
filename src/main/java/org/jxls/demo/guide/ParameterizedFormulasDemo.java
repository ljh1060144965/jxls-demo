package org.jxls.demo.guide;

import org.jxls.common.Context;
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
import java.util.List;
import java.util.Locale;

/**
 * @author Leonid Vysochyn
 */
public class ParameterizedFormulasDemo {
    static Logger logger = LoggerFactory.getLogger(ParameterizedFormulasDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Parameterized Formulas demo");
        List<Employee> employees = generateSampleEmployeeData();
        InputStream is = ParameterizedFormulasDemo.class.getResourceAsStream("param_formulas_template.xls");
        OutputStream os = new FileOutputStream("target/param_formulas_output.xls");
        Context context = new Context();
        context.putVar("employees", employees);
        context.putVar("bonus", 0.1);
        JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
        is.close();
        os.close();
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        return employees;
    }
}
