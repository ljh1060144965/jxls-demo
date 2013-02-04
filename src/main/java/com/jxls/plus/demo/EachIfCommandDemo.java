package com.jxls.plus.demo;

import com.jxls.plus.area.XlsArea;
import com.jxls.plus.command.Command;
import com.jxls.plus.command.EachCommand;
import com.jxls.plus.command.IfCommand;
import com.jxls.plus.common.AreaRef;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Department;
import com.jxls.plus.demo.model.Employee;
import com.jxls.plus.transform.Transformer;
import com.jxls.plus.transform.poi.PoiTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 1/30/12 12:15 PM
 */
public class EachIfCommandDemo {
    static Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String output = "target/each_if_demo_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Each,If command demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        assert is != null;
        logger.info("Creating Workbook");
        Workbook workbook = WorkbookFactory.create(is);
        Transformer poiTransformer = PoiTransformer.createTransformer(workbook);
        System.out.println("Creating area");
        XlsArea xlsArea = new XlsArea("Template!A1:G15", poiTransformer);
        XlsArea departmentArea = new XlsArea("Template!A2:G13", poiTransformer);
        EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
        XlsArea employeeArea = new XlsArea("Template!A9:F9", poiTransformer);
        XlsArea ifArea = new XlsArea("Template!A18:F18", poiTransformer);
        IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
                ifArea,
                new XlsArea("Template!A9:F9", poiTransformer));
        employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
        Command employeeEachCommand = new EachCommand( "employee", "department.staff", employeeArea);
        departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
        xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying at cell " + new CellRef("Down!B2"));
        xlsArea.applyAt(new CellRef("Down!B2"), context);
        xlsArea.processFormulas();
        logger.info("Setting EachCommand direction to Right");
        departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
        logger.info("Applying at cell " + new CellRef("Right!A1"));
        xlsArea.reset();
        xlsArea.applyAt(new CellRef("Right!A1"), context);
        xlsArea.processFormulas();
        logger.info("Complete");
        OutputStream os = new FileOutputStream(output);
        workbook.write(os);
        logger.info("written to file");
        is.close();
        os.close();
    }

    public static List<Department> createDepartments() {
        List<Department> departments = new ArrayList<Department>();
        Department department = new Department("IT");
        Employee chief = new Employee("Derek", 35, 3000, 0.30);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(new Employee("Elsa", 28, 1500, 0.15));
        department.addEmployee(new Employee("Oleg", 32, 2300, 0.25));
        department.addEmployee(new Employee("Neil", 34, 2500, 0.00));
        department.addEmployee(new Employee("Maria", 34, 1700, 0.15));
        department.addEmployee(new Employee("John", 35, 2800, 0.20));
        departments.add(department);
        department = new Department("HR");
        chief = new Employee("Betsy", 37, 2200, 0.30);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(new Employee("Olga", 26, 1400, 0.20));
        department.addEmployee(new Employee("Helen", 30, 2100, 0.10));
        department.addEmployee(new Employee("Keith", 24, 1800, 0.15));
        department.addEmployee(new Employee("Cat", 34, 1900, 0.15));
        departments.add(department);
        department = new Department("BA");
        chief = new Employee("Wendy", 35, 2900, 0.35);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(new Employee("Denise", 30, 2400, 0.20));
        department.addEmployee(new Employee("LeAnn", 32, 2200, 0.15));
        department.addEmployee(new Employee("Natali", 28, 2600, 0.10));
        department.addEmployee(new Employee("Martha", 33, 2150, 0.25));
        departments.add(department);
        return departments;
    }

}