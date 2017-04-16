package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class TestClass {
    private int rowNum = 0;
    
    /**
     * map转换适配器
     * 
     * @param obj
     * @return
     */
    @SuppressWarnings("unchecked")
    public Map<String, Object> ConvertObjToMapAdpater(Object obj) {
        
        // 如果是集合或者数组，获取当前行的数据，最后尝试转换成map
        if (obj instanceof Collection<?>) {
            Collection<?> col = (Collection<?>) obj;
            return this.ConvertObjToMapAdpater(col.toArray()[this.rowNum]);
        } else if (obj instanceof Map) {
            return (Map<String, Object>) obj;
        } else if (obj.getClass().isArray()) {
            return this.ConvertObjToMapAdpater(((Object[]) obj)[this.rowNum]);
        } else {
            return this.ConvertObjToMap(obj);
        }
    }
    
    /**
     * 转换成map
     * 
     * @param obj
     * @return
     */
    public Map<String, Object> ConvertObjToMap(Object obj) {
        Map<String, Object> reMap = new HashMap<String, Object>();
        if (obj == null)
            return null;
        Field[] fields = obj.getClass().getDeclaredFields();
        try {
            for (int i = 0; i < fields.length; i++) {
                try {
                    Field f = obj.getClass().getDeclaredField(
                            fields[i].getName());
                    f.setAccessible(true);
                    Object o = f.get(obj);
                    reMap.put(fields[i].getName(), o);
                } catch (NoSuchFieldException e) {
                    e.printStackTrace();
                } catch (IllegalArgumentException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        } catch (SecurityException e) {
            System.out.println("不支持转换该类型的数据:" + obj.toString());
            e.printStackTrace();
        }
        return reMap;
    }
    
    public void ConvertWord(String fileName, Map<String, Object> map) throws Exception {
        
        String pathFileName = "D:\\201704\\watermark\\testFile\\tempfile_" + System.currentTimeMillis() + ".doc";
        File file = new File(pathFileName); // 创建临时文件
        OutputStream os = new FileOutputStream(file);
        InputStream fileInputStream = new FileInputStream(fileName);

        try {
            // 读取word源文件
            XWPFDocument document = new XWPFDocument(fileInputStream);
            // 替换段落里面的变量
            this.replaceInPara(document, map);
            // 替换表格里面的变量
            this.replaceInTable(document, map);

            document.write(os);
            this.close(os);
            this.close(fileInputStream);

        } catch (Exception e) {
            throw new Exception(e);
        } finally {
            this.close(os);
            this.close(fileInputStream);
        }

    }

    /**
     * 替换段落里面的变量
     * 
     * @param doc
     *            要替换的文档
     * @param params
     *            参数
     */
    private void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            Map<String, Object> map = getParagraphData(para, params);
            this.replaceInPara(para, map);
        }
    }

    /**
     * 替换段落里面的变量
     * 
     * @param para
     *            要替换的段落
     * @param params
     *            参数
     */
    private void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        List<XWPFRun> runs;

        String runText = "";
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            if (runs.size() > 0) {
                int j = runs.size();
                for (int i = 0; i < j; i++) {
                    XWPFRun run = runs.get(0);
                    String i1 = run.toString();
                    runText += i1;
                    para.removeRun(0);
                }

                // 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                // 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                runText = this.replaceText(runText, params);
                para.insertNewRun(0).setText(runText);
            }
        }
    }

    /**
     * 替换字符串中所有的匹配项
     * 
     * @param text
     * @param params
     * @return
     */
    private String replaceText(String text, Map<String, Object> params) {
        Matcher matcher = this.matcher(text);
        String key = "";
        
        if (matcher.find()) {
            key = matcher.group(1);
            if (params.containsKey(key)) {
                text = matcher.replaceFirst(String.valueOf(params.get(key)));
            } else {
                text = matcher.replaceFirst("");
            }
        }

        return text;
    }
    
    /**
     * 替换表格里面的变量
     * 
     * @param doc
     *            要替换的文档
     * @param params
     *            参数
     */
    private void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        
        int modelRowCount = 0;
        List<XWPFTableRow> rows;

        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            modelRowCount = rows.size();
            for (int i = (modelRowCount - 1); i >= 0; i--) {
                XWPFTableRow dataRow = rows.get(i);
                int rowCount = this.getRowCount(dataRow, params);
                if (rowCount == 0) {
                    Map<String, Object> rowData = this.getRowData(dataRow, params);
                    this.replaceInRow(table, dataRow, rowData);
                    continue;
                }
                
                for (int j = 0; j < rowCount; j++) {
                    this.copyTableRow(table, dataRow, (i + j), params);
                    dataRow = table.getRows().get((i + j));
                    
                    if (j == (rowCount - 1)) {
                        table.removeRow(i);
                    }
                }
            }
            
            this.rowNum = 0;
        }
    }
    
    private void replaceInRow(XWPFTable table, XWPFTableRow newRow, Map<String, Object> params) {
        for (XWPFTableCell cell : newRow.getTableCells()) {
            this.replaceInCell(cell, params);
        }
    }
    
    private void replaceInCell(XWPFTableCell cell, Map<String, Object> params) {
        for (XWPFParagraph para : cell.getParagraphs()) {
            this.replaceInPara(para, params);
        }
    }
    
    /**
     * 获取当前行数据map
     * 
     * @param row
     * @param params
     * @return
     */
    private Map<String, Object> getRowData(XWPFTableRow row, Map<String, Object> params) {
        List<XWPFParagraph> paras;
        List<XWPFTableCell> cells = row.getTableCells();
        Map<String, Object> map = new HashMap<String, Object>();
        
        for (XWPFTableCell cell : cells) {
            paras = cell.getParagraphs();
            for (XWPFParagraph para : paras) {
                map.putAll(this.getParagraphData(para, params));
            }
        }
        
        return map;
    }
    
    /**
     * 获取当前字符串中的数据Map
     * 
     * @param para
     * @param params
     * @return
     */
    private Map<String, Object> getParagraphData(XWPFParagraph para, Map<String, Object> params) {
        Map<String, Object> map = new HashMap<String, Object>();
        Map<String, Object> data = params;
        Matcher matcher = this.matcher(para.getParagraphText());
        if (matcher.find()) {
            String key = matcher.group(1);
            String[] keys = key.split("\\.");
            for (int i = 0; i < keys.length; i++) {
                Object val = data.get(keys[i]);
                if (val == null) {
                    map.put(key, "");
                    break;
                }
                
                if (i != (keys.length - 1)) {
                    data = this.ConvertObjToMapAdpater(val);
                } else {
                    map.put(key, String.valueOf(val));
                    data = params;
                }
            }
        }
        
        return map;
    }

    /**
     * 复制表格的行
     * 
     * @param table
     * @param row
     * @param rowData 
     * @param pos
     * @param params 
     */
    private void copyTableRow(XWPFTable table, XWPFTableRow row, int pos, Map<String, Object> params) {
        XWPFTableRow newRow = new XWPFTableRow(row.getCtRow(), table);
        if (table.addRow(newRow, pos)) {
            this.rowNum++;
            Map<String, Object> rowData = this.getRowData(row, params);
            this.replaceInRow(table, newRow, rowData);
        }
    }
    
    /**
     * 获取表格数据项的大小
     * 
     * @param row
     * @param params
     * @return
     */
    private int getRowCount(XWPFTableRow row, Map<String, Object> params) {
        int rowCount = 0;
        Matcher matcher = null;
        String[] key;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        
        cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            
            paras = cell.getParagraphs();
            for (XWPFParagraph para : paras) {
                
                matcher = this.matcher(para.getParagraphText());
                if (matcher.find()) {
                    key = (matcher.group(1)).split("\\.");
                    if (params.get(key[0]) instanceof Collection<?>) {
                        Collection<?> col = (Collection<?>) params.get(key[0]);
                        rowCount = col.size();
                    } else if (params.get(key[0]).getClass().isArray()) {
                        Object[] objs = (Object[]) params.get(key[0]);
                        rowCount = objs.length;
                    } else {
                        return rowCount;
                    }
                    
                    if (rowCount > 0) {
                        return rowCount;
                    }
                }
            }
        }
        
        return rowCount;
    }

    /**
     * 正则匹配字符串
     * 
     * @param str
     * @return
     */
    public Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 关闭输入流
     * 
     * @param is
     */
    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     * 
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {

        try {
            long startTime = System.currentTimeMillis();
            System.out.println("start:" + startTime);
            Map<String, Object> map = new HashMap<String, Object>();
            
            Map<String, Object> map1 = new HashMap<String, Object>();
            map1.put("idCard", "440104195404242811");
            map1.put("name", "测试人员");
            map1.put("sex", "男");
            map1.put("date", "2017-04-10");
            map1.put("authCode", "152452");
            
            map.put("info", map1);

            List<Map<String, Object>> maps = new ArrayList<Map<String, Object>>();
            Map<String, Object> maps1 = new HashMap<String, Object>();
            maps1.put("cell1", "idcard-1");
            maps1.put("cell2", "name-1");
            maps1.put("cell3", "sex-1");
            maps1.put("cell4", "123-1");
            maps1.put("cell5", "123-1");
            maps1.put("cell6", "123-1");
            maps1.put("cell7", "123-1");
            maps1.put("cell8", "123-1");
            maps1.put("cell9", "123-1");
            maps1.put("cell10", "123-1");
            
            Map<String, Object> maps2 = new HashMap<String, Object>();
            maps2.put("cell1", "idcard-2");
            maps2.put("cell2", "name-2");
            maps2.put("cell3", "sex-2");
            maps2.put("cell4", "123-2");
            maps2.put("cell5", "123-2");
            maps2.put("cell6", "123-2");
            maps2.put("cell7", "123-2");
            maps2.put("cell8", "123-2");
            maps2.put("cell9", "123-2");
            maps2.put("cell10", "123-2");
            
            Map<String, Object> maps3 = new HashMap<String, Object>();
            maps3.put("cell1", "idcard-2");
            maps3.put("cell2", "name-2");
            maps3.put("cell3", "sex-2");
            maps3.put("cell4", "123-2");
            maps3.put("cell5", "123-2");
            maps3.put("cell6", "123-2");
            maps3.put("cell7", "123-2");
            maps3.put("cell8", "123-2");
            maps3.put("cell9", "123-2");
            maps3.put("cell10", "123-2");

            maps.add(maps1);
            maps.add(maps2);
            maps.add(maps3);

            map.put("list", maps);

            TestClass testClass = new TestClass();
            testClass.ConvertWord("d:/project/test/src/test/person_pay_history_three_insurance.docx", map);
            long endTime = System.currentTimeMillis();
            System.out.println("end:" + endTime);
            System.out.println("useTime:" + (endTime - startTime));
            System.out.println("success!1:1100--2:");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
