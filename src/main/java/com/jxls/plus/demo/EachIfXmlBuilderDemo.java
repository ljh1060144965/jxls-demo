package com.jxls.plus.demo;

import com.jxls.plus.area.Area;
import com.jxls.plus.builder.AreaBuilder;
import com.jxls.plus.builder.xml.XmlAreaBuilder;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.demo.model.Department;
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
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/14/12 3:59 PM
 */
public class EachIfXmlBuilderDemo {
    static Logger logger = LoggerFactory.getLogger(EachIfCommandDemo.class);
    private static String template = "each_if_demo.xls";
    private static String xmlConfig = "each_if_demo.xml";
    private static String output = "target/each_if_xml_builder_output.xls";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        logger.info("Executing Each,If XML builder demo");
        execute();
    }

    public static void execute() throws IOException, InvalidFormatException {
        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        InputStream is = EachIfCommandDemo.class.getResourceAsStream(template);
        OutputStream os = new FileOutputStream(output);
        Transformer transformer = PoiTransformer.createTransformer(is, os);
        System.out.println("Creating areas");
        InputStream configInputStream = EachIfXmlBuilderDemo.class.getResourceAsStream(xmlConfig);
        AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Area xlsArea2 = xlsAreaList.get(1);
        Context context = new Context();
        context.putVar("departments", departments);
        logger.info("Applying first area at cell " + new CellRef("Down!A1"));
        xlsArea.applyAt(new CellRef("Down!A1"), context);
        xlsArea.processFormulas();
        logger.info("Applying second area at cell " + new CellRef("Right!A1"));
        xlsArea.reset();
        xlsArea2.applyAt(new CellRef("Right!A1"), context);
        xlsArea2.processFormulas();
        logger.info("Complete");
        transformer.write();
        logger.info("written to file");
        is.close();
        os.close();
    }

}
