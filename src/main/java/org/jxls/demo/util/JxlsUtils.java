package org.jxls.demo.util;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.CellRefUtil;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel表格生成工具类（Jxls2.x模板方式生成）
 *
 * jxls官方文档：
 * http://jxls.sourceforge.net/samples/object_collection.html
 * 官方demo：
 * https://bitbucket.org/leonate/jxls-demo
 *
 * @author : lijiahui
 * @version : 1.0
 * @date : 2020/11/4 13:12
 */
public class JxlsUtils {

    private static final Logger LOG = LoggerFactory.getLogger(JxlsUtils.class);

    private static final String DELIMITER1 = "!";
    private static final String DELIMITER2 = ":";
    /** 表单第一个单元格位置 **/
    private static final String A1 = "A1";

    /**
     * excel生成（用于直接生成excel文件，不需要二次加工处理Workbook）
     *
     * @param templatePath
     * @param destFilePath
     * @param beanParams 模型参数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void buildExcel(String templatePath, String destFilePath, Map<String, Object> beanParams)
            throws IOException, InvalidFormatException {
        File template = getTemplate(templatePath);

        try (final InputStream is = new FileInputStream(template)) {
            buildExcel(is, destFilePath, beanParams);
        }
    }

    /**
     * excel生成（用于直接生成excel文件，不需要二次加工处理Workbook）
     *
     * @param is
     * @param destFilePath
     * @param beanParams 模型参数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void buildExcel(InputStream is, String destFilePath, Map<String, Object> beanParams)
            throws IOException, InvalidFormatException {
        try (final OutputStream os = new FileOutputStream(destFilePath)) {
            buildExcel(is, os, beanParams);
        }
    }

    /**
     * excel生成（用于直接生成excel文件，不需要二次加工处理Workbook）
     *
     * @param is
     * @param os
     * @param beanParams 模型参数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void buildExcel(InputStream is, OutputStream os, Map<String, Object> beanParams)
            throws IOException, InvalidFormatException {
        Context context = getContext(beanParams);

        PoiTransformer transformer = PoiTransformer.createTransformer(is, os);
        // with multi sheets it is better to use StandardFormulaProcessor by disabling the FastFormulaProcessor
        JxlsHelper.getInstance().setUseFastFormulaProcessor(false)
                .processTemplate(context, transformer);
    }

    /**
     * excel生成（带有Workbook返回，用于需要二次加工处理，且不用先去生成excel文件）
     *
     * 注：此方法未生成excel文件，只生成了Workbook对象,
     * 用法与jxls1.x中net.sf.jxls.transformer.XLSTransformer#transformXLS(java.io.InputStream, java.util.Map)一致
     *
     * @param templatePath
     * @param beanParams 模型参数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static Workbook buildExcel(String templatePath, Map<String, Object> beanParams)
            throws IOException, InvalidFormatException {
        File template = getTemplate(templatePath);
        try (final InputStream is = new FileInputStream(template)) {
            return buildExcel(is, beanParams);
        }
    }

    /**
     * excel生成（带有Workbook返回，用于需要二次加工处理，且不用先去生成excel文件）
     *
     * 注：此方法未生成excel文件，只生成了Workbook对象,
     * 用法与jxls1.x中net.sf.jxls.transformer.XLSTransformer#transformXLS(java.io.InputStream, java.util.Map)一致
     *
     * @param is
     * @param beanParams 模型参数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static Workbook buildExcel(InputStream is, Map<String, Object> beanParams)
            throws IOException, InvalidFormatException {
        Context context = getContext(beanParams);
        PoiTransformer transformer = PoiTransformer.createTransformer(is);

        return processTemplate(context, transformer);
    }

    public static Workbook processTemplate(Context context, PoiTransformer transformer) {
        JxlsHelper jxlsHelper = JxlsHelper.getInstance().setUseFastFormulaProcessor(false);
        AreaBuilder areaBuilder = jxlsHelper.getAreaBuilder();
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        // 修复非首次构建无效问题
        // 由于当前使用的是基于excel注释模式构建的，每次构建完，注释就会被清除，后续再次构建时就会无注释，导致构建无效，
        if (CollectionUtils.isEmpty(xlsAreaList)) {
            XlsArea xlsArea = getXlsArea(transformer);
            xlsAreaList.add(xlsArea);
        }

        for (Area xlsArea : xlsAreaList) {
            xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
        }

        if (jxlsHelper.isProcessFormulas()) {
            for (Area xlsArea : xlsAreaList) {
                FormulaProcessor fp = jxlsHelper.getFormulaProcessor();
                if (fp == null) {
                    if (jxlsHelper.isUseFastFormulaProcessor()) {
                        fp = new FastFormulaProcessor();
                    } else {
                        fp = new StandardFormulaProcessor();
                    }
                }
                xlsArea.setFormulaProcessor(fp);
                xlsArea.processFormulas();
            }
        }

        return transformer.getWorkbook();
    }

    private static XlsArea getXlsArea(PoiTransformer transformer) {
        Sheet sheet = transformer.getWorkbook().getSheetAt(0);
        // 表格区域的最后一行行号，从0开始
        int lastRowNum = sheet.getLastRowNum();
        // 表格区域的最后一列列号，从0开始(要注意下有时候area写多了，这里获取的列会变成-1，可以后续判断小于0，则赋值个50)
        int lastCellNum = sheet.getRow(lastRowNum).getLastCellNum() - 1;
        String sheetName = sheet.getSheetName();

        if (lastCellNum < 0) {
            lastCellNum = 50;
        }

        // 如：Template!A1:D4
        String areaRef = StringUtils.join(sheetName, DELIMITER1, A1, DELIMITER2,
                getCellReference(lastCellNum, lastRowNum));
        return new XlsArea(areaRef, transformer);
    }

    /**
     * 获取单元格位置（例如：输入 0,0，输出A1）
     *
     * @param column 从0开始
     * @param row 从0开始
     * @return
     */
    public static String getCellReference(int column, int row) {
        int realRow = row + 1;
        return StringUtils.join(CellRefUtil.convertNumToColString(column), realRow);
    }

    private static File getTemplate(String path) {
        File template = new File(path);
        if(!template.exists()){
            throw new RuntimeException(
                    MessageFormat.format("Excel模板【{0}】未找到", path));
        }
        return template;
    }

    private static Context getContext(Map<String, Object> beanParams) {
        Context context = PoiTransformer.createInitialContext();
        if (MapUtils.isNotEmpty(beanParams)) {
            beanParams.forEach((key, value) ->  context.putVar(key, beanParams.get(key)));
        }
        return context;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        String template = "/demo_temp.xls";
        String output = "target/demo_output.xls";
        Map<String, Object> map = new HashMap<>(16);
        map.put("list", Arrays.asList(1, 2, 3, 4, 5));
        Workbook wb = JxlsUtils.buildExcel(template, map);
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(0);
        row.getCell(0).setCellValue("xxxxxx6666666");
        wb.write(new FileOutputStream(output));
    }
}
