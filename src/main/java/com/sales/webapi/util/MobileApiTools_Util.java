package com.sales.webapi.util;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.sales.webapi.util.Excel_Util;

public class MobileApiTools_Util {
    private MobileApiTools_Util() {
    }

    /**
     * 获取当前系统时间
     * 
     * @return
     */
    public static String getDate() {
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return df.format(new Date());
    }

    /**
     * 期望结果与实际结果比较
     * 
     * @param expectedResult
     * @param actualResult
     * @return
     */
    public static String assertResult(String expectedResult, String actualResult) {
        String result;
        if (expectedResult.equals(actualResult))
            result = "OK";
        else
            result = "NG";
        return result;

    }

    /**
     * 初始化 Excel 中指定列数据
     * 
     * @param filePath
     *            文件路径
     * @param sheetName
     *            sheet 名称
     * @param titleLineIndex
     *            表标题所在行索引
     * @param columnName
     *            列名称
     * @param content
     *            需写入的内容
     * @param color
     *            需设置的单元格颜色
     * @throws IOException
     */
    public static void initializeData(int titleLineIndex, String columnName,
            String content, int color) throws IOException {
        int maxRow = Excel_Util.getInstance().getSheetMaxRow();
        int columnIndex = Excel_Util.getInstance().getColumnIndex(
                titleLineIndex, columnName);
        if (maxRow != -1 && columnIndex != -1) {
            for (int i = titleLineIndex + 1; i <= maxRow; i++) {
                // 初始化单元格内容
                Excel_Util.getInstance().writeExcelCell(i, columnIndex, content);
                // 设置单元格颜色
                Excel_Util.getInstance().setCellBackgroundColor(titleLineIndex,
                        columnName, i, color);
            }
        }
    }

    /**
     * 将执行结果写入 Excel 中
     * 
     * @param filePath
     *            文件路径
     * @param sheetName
     *            sheet 名称
     * @param titleLineIndex
     *            标题所在索引行
     * @param writeStartRow
     *            写入起始行
     * @param columnName
     *            列名称
     * @param content
     *            写入内容
     */
    public static void writeResult(int titleLineIndex, int writeStartRow,
            String columnName, String content) {

        int columnIndex = Excel_Util.getInstance().getColumnIndex(
                titleLineIndex, columnName);
        if (columnIndex != -1) {
            try {
                Excel_Util.getInstance().writeExcelCell(writeStartRow,
                        columnIndex, content);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 表标题行是否包含指定参数
     * @param argLineIndex            参数名称所在单元格行索引
     * @param argColumnIndex        参数名称所在单元格列索引
     * @param titleLineIndex        表标题所在行索引
     * @return
     */
    public static boolean isArgEquals(int argLineIndex, int argColumnIndex,
            int titleLineIndex) {
        String[] argArray = null;
        boolean flag = false;
        
        // 获取所在单元格的参数列表
        String args = Excel_Util.getInstance().readExcelCell(argLineIndex, argColumnIndex);
        if (!args.equals("")) {
            argArray = args.split("\\|");
        } else {
            System.out.println("文件: " + Excel_Util.getInstance().getFilePath()
                    + ", sheetName: " + Excel_Util.getInstance().getSheetName()
                    + ", 获取参数失败...");
        }
        
        if (argArray != null) {
            for (int i = 0; i < argArray.length; i++) {
                int tempIndex = Excel_Util.getInstance().getColumnIndex(
                        titleLineIndex, argArray[i]);
                if (tempIndex >= 0)
                    flag = true;
                else
                    System.out.println("参数: " + argArray[i]
                            + ", 不存在表标题行中, 请检查...");
            }
        }
        return flag;
    }

    public void info() {
        
    }
}