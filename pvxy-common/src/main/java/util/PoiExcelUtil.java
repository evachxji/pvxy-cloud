package util;

import cn.hutool.core.util.URLUtil;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.net.URI;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * POI解析EXCEL工具类
 *
 * @author yejuncheng
 * @date 2020/5/7 13:29
 */
@Getter
@Setter
@Slf4j
public class PoiExcelUtil {

    private Workbook workbook;

    private CellStyle cellStyle;

    private Sheet sheet;

    private int currentRowNum = 0;

    private int code = 500;

    private PoiExcelUtil() {
        workbook = new XSSFWorkbook();
        cellStyle = workbook.createCellStyle();
        sheet = workbook.createSheet();
        init();
    }

    private PoiExcelUtil(InputStream in) {
        try {
            workbook = WorkbookFactory.create(in);
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }
        cellStyle = workbook.createCellStyle();
        sheet = workbook.getSheetAt(0);
        init();
    }

    /**
     * 创建空白Excel
     */
    public static PoiExcelUtil newExcel() {
        return new PoiExcelUtil();
    }

    /**
     * 加载本地文件 (Path)
     */
    public static PoiExcelUtil loadPath(String path) {
        return loadFile(new File(path));
    }

    /**
     * 加载本地文件 (File)
     */
    public static PoiExcelUtil loadFile(File file) {
        InputStream in = null;
        try {
            in = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        }
        return new PoiExcelUtil(in);
    }

    /**
     * 加载本地文件 (MultipartFile)
     */
    public static PoiExcelUtil loadFile(MultipartFile file) {
        InputStream in = null;
        try {
            in = file.getInputStream();
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
        return new PoiExcelUtil(in);
    }

    /**
     * 加载url文件
     */
    public static PoiExcelUtil loadUrl(URI url) {
        File file = new File(url);
        return loadFile(file);
    }

    public boolean loadSuccess() {
        return this.code == 200 && workbook != null;
    }

    public String read(int cellnum) {
        return getValue(this.sheet.getRow(this.currentRowNum).getCell(cellnum));
    }

    public String readDate(int cellnum) {
        return getDateValue(this.sheet.getRow(this.currentRowNum).getCell(cellnum));
    }

    public String read(int rownow, int cellnum) {
        return getValue(this.sheet.getRow(rownow).getCell(cellnum));
    }

    public String readDate(int rownow, int cellnum) {
        return getDateValue(this.sheet.getRow(rownow).getCell(cellnum));
    }

    /**
     * 默认当前行，指定列号 write
     *
     * @param cellnum
     * @param object
     * @return
     * @throws Exception
     */
    public PoiExcelUtil write(int cellnum, Object object) {
        if (object == null) {
            return this;
        }
        return write(currentRowNum, cellnum, object);
    }

    /**
     * 默认当前行，从第一列开始write
     *
     * @return
     * @throws Exception
     */
    public PoiExcelUtil write(Object... objects) {
        for (int i = 0; i < objects.length; i++) {
            Object obj = objects[i];
            if (obj == null) {
                continue;
            }
            write(currentRowNum, i, obj);
        }
        return this;
    }

    /**
     * write到第一行，作标题
     *
     * @return
     * @throws Exception
     */
    public PoiExcelUtil title(Object... objects) {
        for (int i = 0; i < objects.length; i++) {
            Object obj = objects[i];
            if (obj == null) {
                continue;
            }
            write(0, i, obj);
        }

        int[] widths = new int[objects.length];
        for (int i = 0; i < objects.length; i++) {
            widths[i] = objects[i].toString().length() * 3;
        }
        setWidth(widths);
        return this;
    }

    public PoiExcelUtil write(int rownum, int cellnum, Object object) {
        if (object == null) {
            return this;
        }
        Row row = this.sheet.getRow(rownum);
        if (row == null) {
            row = this.sheet.createRow(rownum);
        }
        Cell cell = row.getCell(cellnum);
        if (cell == null) {
            cell = row.createCell(cellnum);
        }
        //水平居中
        cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
        cell.getCellStyle().setWrapText(true);

        if (object instanceof Double) {
            double temp = (double) object;
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Float) {
            double temp = getDouble((float) object, 2);
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Integer) {
            cell.setCellValue((int) object);
        } else if (object instanceof BigDecimal) {
            double temp = ((BigDecimal) object).doubleValue();
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Long) {
            cell.setCellValue((long) object);
        } else if (object instanceof Date) {
            cell.setCellValue((Date) object);
        } else if (object instanceof Boolean) {
            cell.setCellValue((boolean) object);
        } else if (object instanceof Character) {
            cell.setCellValue((Character) object);
        } else if (object instanceof String) {
            String str = (String) object;
            cell.setCellValue(str);
        } else if (object instanceof LocalDate) {
            cell.setCellValue(((LocalDate) object).format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
        } else if (object instanceof LocalDateTime) {
            cell.setCellValue(((LocalDateTime) object).format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        } else {
            log.error("未知的数据类型！");
        }
        return this;
    }

    public PoiExcelUtil writeLeft(int rownum, int cellnum, Object object) {
        if (object == null) {
            return this;
        }
        Row row = this.sheet.getRow(rownum);
        if (row == null) {
            row = this.sheet.createRow(rownum);
        }
        Cell cell = row.getCell(cellnum);
        if (cell == null) {
            cell = row.createCell(cellnum);
        }
        //水平居中
        cell.getCellStyle().setAlignment(HorizontalAlignment.LEFT);

        if (object instanceof Double) {
            double temp = (double) object;
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Float) {
            double temp = getDouble((float) object, 2);
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Integer) {
            cell.setCellValue((int) object);
        } else if (object instanceof BigDecimal) {
            double temp = ((BigDecimal) object).doubleValue();
            String str = String.valueOf(temp);
            cell.setCellValue(removeZero(str));
        } else if (object instanceof Long) {
            cell.setCellValue((long) object);
        } else if (object instanceof Date) {
            cell.setCellValue((Date) object);
        } else if (object instanceof Boolean) {
            cell.setCellValue((boolean) object);
        } else if (object instanceof Character) {
            cell.setCellValue((Character) object);
        } else if (object instanceof String) {
            String str = (String) object;
            cell.setCellValue(removeZero(str));
        } else if (object instanceof LocalDate) {
            cell.setCellValue(((LocalDate) object).format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
        } else if (object instanceof LocalDateTime) {
            cell.setCellValue(((LocalDateTime) object).format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        } else {
            log.error("未知的数据类型！");
        }
        return this;
    }

    public PoiExcelUtil setBlank(int rownum, int cellnum) {
        Row row = this.sheet.getRow(rownum);
        if (row == null) {
            row = this.sheet.createRow(rownum);
        }
        Cell cell = row.getCell(cellnum);
        if (cell == null) {
            cell = row.createCell(cellnum);
        }
        cell.setBlank();
        return this;
    }

    public PoiExcelUtil setRowHeight(short height) {
        Row row = this.sheet.getRow(this.getCurrentRowNum());
        if (row == null) {
            row = this.sheet.createRow(this.getCurrentRowNum());
        }
        row.setHeight(height);
        return this;
    }

    public int getRowNum() {
        return sheet.getPhysicalNumberOfRows();
    }

    public PoiExcelUtil nextRow() {
        this.currentRowNum++;
        return this;
    }

    public boolean hasNextRow() {
        return this.sheet.getRow(currentRowNum + 1) != null;
    }

    public PoiExcelUtil nextSheet() {
        this.sheet = workbook.getSheetAt(1);
        this.setCurrentRowNum(0);
        return this;
    }

    /**
     * 向下跨越n行，输入1等于不传参
     *
     * @param num
     * @return
     */
    public PoiExcelUtil nextRow(int num) {
        this.currentRowNum += num;
        return this;
    }

    /**
     * 调用该方法后，主方法不要return或者return null，否则后端会报错：
     * InvalidMimeTypeException: Invalid mime type "bin; charset=UTF-8": does not contain '/'
     *
     * @param fileName
     * @param response
     * @return
     */
    public HttpServletResponse export(String fileName, HttpServletResponse response) {
        //设置文件名
        response.setContentType("bin");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLUtil.encode(fileName) + ".xlsx");
        //下载输出流
        try {
            workbook.write(response.getOutputStream());
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return response;
    }

    /**
     * 导出到本地目录，不能导出到C盘根目录下，必须要有2个层级目录
     *
     * @param filePath
     * @return
     */
    public File export(String filePath) {
        filePath += ".xlsx";
        File file = new File(filePath);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        if (!file.exists()) {
            try {
                boolean success = file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(filePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        }
        //下载输出流
        try {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return file;
    }

    private static String getValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();
                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                try {
                    returnValue = String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    returnValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    private static String getDateValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                short format = cell.getCellStyle().getDataFormat();
                if (DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = null;
//                    System.out.println("cell.getCellStyle().getDataFormat()=" + cell.getCellStyle().getDataFormat());
                    if (format == 20 || format == 32) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else if (format == 14 || format == 31 || format == 57 || format == 58) {
                        // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = DateUtil
                                .getJavaDate(value);
                        returnValue = sdf.format(date);
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }
                    try {
                        returnValue = sdf.format(cell.getDateCellValue());
                    } catch (Exception e) {
                        try {
                            throw new Exception("exception on get date data !".concat(e.toString()));
                        } catch (Exception e1) {
                            e1.printStackTrace();
                        }
                    }
                } else {
                    BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
                    returnValue = bd.toPlainString();// 数值 这种用BigDecimal包装再获取plainString，可以防止获取到科学计数值
                }
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                try {
                    returnValue = String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    returnValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }


    private static Double getDouble(Cell cell) {
        if (cell == null) {
            return 0d;
        }
        return cell.getNumericCellValue();
    }

    /**
     * 将InputStream写入本地文件
     *
     * @param input 输入流
     * @throws IOException IOException
     */
    public static void writeToLocal(String fileName, InputStream input) throws IOException {
        int index;
        byte[] bytes = new byte[1024];
        FileOutputStream downloadFile = new FileOutputStream(fileName);
        while ((index = input.read(bytes)) != -1) {
            downloadFile.write(bytes, 0, index);
            downloadFile.flush();
        }
        input.close();
        downloadFile.close();
    }

    /**
     * 设置列宽
     *
     * @param index 列号
     * @param count 宽度占几个字符
     * @return
     */
    public PoiExcelUtil setWidth(int index, int count) {
        this.sheet.setColumnWidth(index, count * 256);
        return this;
    }

    /**
     * 设置列宽
     *
     * @return
     */
    public PoiExcelUtil setWidth(int... counts) {
        for (int i = 0; i < counts.length; i++) {
            this.sheet.setColumnWidth(i, counts[i] * 256);
        }
        return this;
    }

    /**
     * 合并单元格
     *
     * @param sX 开始行
     * @param eX 结束行
     * @param sY 开始列
     * @param eY 结束列
     * @return
     */
    public boolean merge(int sX, int eX, int sY, int eY) {
        CellRangeAddress region = new CellRangeAddress(sX, eX, sY, eY);
        sheet.addMergedRegion(region);
        return true;
    }

    /**
     * 合并单元格(当前行,向下合并)
     *
     * @param height 向下合并几格
     * @param ys 第几列(多个)
     * @return
     */
    public PoiExcelUtil mergeY(int height,int... ys) {
        if (height <= 1) {
            return this;
        }
        for (int y : ys) {
            CellRangeAddress region = new CellRangeAddress(currentRowNum, currentRowNum + height - 1, y, y);
            sheet.addMergedRegion(region);
        }
        return this;
    }


    /**
     * 指定index插入行（本质上是将所有其他行下移，再插入到当前行）
     */
    public PoiExcelUtil insertRows() {
        return insertRows(currentRowNum);
    }

    /**
     * 指定index插入行（本质上是将所有其他行下移，再插入到当前行）
     */
    public PoiExcelUtil insertRows(int index) {
        if (sheet.getRow(index) != null) {
            int lastRowNo = sheet.getLastRowNum();
            sheet.shiftRows(index, lastRowNo, 1);
        }
        sheet.createRow(index);
        return this;
    }


    private void init() {
        // 公式自动计算
        sheet.setForceFormulaRecalculation(true);
        // 自动换行
        cellStyle.setWrapText(true);
        // 水平左对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 竖直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        code = 200;
    }

    /**
     * 保留 n 位小数
     *
     * @param num
     * @param n
     * @return
     */
    public static double getDouble(double num, int n) {
        double rate = Math.pow(10d, (double) n);
        return (double) Math.round(num * rate) / rate;
    }

    /**
     * 去掉小数点后面多余的0
     *
     * @param str
     * @return
     */
    public static String removeZero(String str) {
        if (str.contains(".") && str.endsWith("0")) {
            for (int i = str.length() - 1; i > 0; i--) {
                if (str.charAt(i) != '0') {
                    if (str.charAt(i) == '.') {
                        str = str.substring(0, i);
                    } else {
                        str = str.substring(0, i + 1);
                    }
                    break;
                }
            }
        }
        return str;
    }
}
