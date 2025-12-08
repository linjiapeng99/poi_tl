package com.example.poi_demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.data.FilePictureRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.util.TableTools;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.util.*;

@SpringBootTest
class PoiDemoApplicationTests {

    @Test
    void contextLoads() {
    }

    /**
     * poi测试列表的渲染
     *
     * @throws IOException
     */
    @Test
    void testRenderTemplate1() throws IOException {
        List<Person> personList = Arrays.asList(
                new Person("张三", 22, "男", "NBKJ", "无"),
                new Person("李四", 23, "女", "NBKJ", "无")
        );
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        String templatePath = "src/main/resources/templateDoc/poi测试.docx";
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        Configure configure = builder.build();
        configure.customPolicy("personList", policy);
        XWPFTemplate template = XWPFTemplate.compile(templatePath, configure).render(
                new HashMap<String, Object>() {{
                    put("personList", personList);
                }}
        );
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/output.docx");
        template.write(out);
        template.close();
        out.close();
    }

    /**
     * poi测试表头复杂的列表1
     */
    @Test
    void testRenderTemplate2() throws IOException {
        //1. 准备数据
        List<ServEnv> servEnvList = Arrays.asList(
                // 服务器1（id=1）的两个软件
                new ServEnv(1, "Windows系统", "64*1T", "Mysql", "V1.0", "NBKJ", "存储数据"),
                new ServEnv(1, "Windows系统", "64*1T", "Redis", "V2.0", "NBKJ", "缓存数据"),
                new ServEnv(1, "Windows系统", "64*1T", "MQ", "V3.0", "NBKJ", "消息队列"),
                new ServEnv(1, "Linux系统", "32*512G", "Tomcat", "V9.0", "Apache", "Web容器"),
                new ServEnv(1, "Linux系统", "32*512G", "Nginx", "V1.18", "开源", "反向代理")
        );
        //2.配置绑定
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        //3.创建表格模板
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy() {
            @Override
            protected void afterloop(XWPFTable table, Object data) {
                // 先调用父类方法完成数据渲染
                super.afterloop(table, data);

                // 根据ID合并单元格
                mergeCellsById(table);
            }

            /**
             * 核心合并方法：根据ID列的值合并相同ID的行
             * 表格结构：
             * - 行0: 主表头
             * - 行1: 子表头
             * - 行2+: 数据行
             * 列结构：
             * - 列0: ID（序号）
             * - 列1: 硬件名称
             * - 列2: 硬件配置
             * - 列3-5: 软件信息（不合并）
             */
            private void mergeCellsById(XWPFTable table) {
                // 数据行起始位置（跳过2行表头）
                final int DATA_START_ROW = 2;
                int totalRows = table.getNumberOfRows();

                // 没有足够数据行时直接返回
                if (totalRows <= DATA_START_ROW) return;

                // 从第一个数据行开始扫描
                int currentRow = DATA_START_ROW;

                while (currentRow < totalRows) {
                    // 获取当前行的ID值
                    String currentId = getCellText(table, currentRow, 0);
                    if (currentId.isEmpty()) {
                        currentRow++;
                        continue;
                    }

                    // 查找相同ID的结束行
                    int endRow = currentRow;
                    for (int nextRow = currentRow + 1; nextRow < totalRows; nextRow++) {
                        String nextId = getCellText(table, nextRow, 0);
                        if (currentId.equals(nextId)) {
                            endRow = nextRow;
                        } else {
                            break;
                        }
                    }

                    // 如果有多行相同ID，合并这三列
                    if (endRow > currentRow) {
                        // 合并ID列（列0）
                        TableTools.mergeCellsVertically(table, 0, currentRow, endRow);
                        // 合并硬件名称列（列1）
                        TableTools.mergeCellsVertically(table, 1, currentRow, endRow);
                        // 合并硬件配置列（列2）
                        TableTools.mergeCellsVertically(table, 2, currentRow, endRow);

                        // 跳转到下一组数据的开始
                        currentRow = endRow + 1;
                    } else {
                        // 单行数据，跳到下一行
                        currentRow++;
                    }
                }
            }

            /**
             * 安全获取单元格文本
             */
            private String getCellText(XWPFTable table, int row, int col) {
                try {
                    XWPFTableCell cell = table.getRow(row).getCell(col);
                    return cell.getText() != null ? cell.getText().trim() : "";
                } catch (Exception e) {
                    return "";
                }
            }
        };
        //模板路径
        String templatePath = "src/main/resources/templateDoc/环境配置1.docx";
        Configure configure = builder.build();
        configure.customPolicy("servEnvList", policy);
        XWPFTemplate template = XWPFTemplate.compile(templatePath, configure).render(
                new HashMap<String, Object>() {{
                    put("servEnvList", servEnvList);
                }}
        );
        //4.输出文件
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/输出环境配置4.docx");
        template.write(out);
        template.close();
        out.close();
    }

    /**
     * poi测试表头复杂的列表2
     */
    @Test
    void testRenderTemplate3() throws IOException {
        //1. 准备数据
        List<Software> swList = Arrays.asList(
                new Software("Mysql", "V1.0", "NBKJ", "存储数据"),
                new Software("Redis", "V2.0", "NBKJ", "缓存数据"),
                new Software("MQ", "V3.0", "NBKJ", "消息队列")
        );
        List<ServEnv> servEnvList = Arrays.asList(
                new ServEnv(1, "Windows服务器", "64*1T", swList)
        );
        //2.配置绑定
        ConfigureBuilder builder = Configure.builder();
        //3.创建表格模板
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        //模板路径
        String templatePath = "src/main/resources/templateDoc/环境配置.docx";
        Configure configure = builder.build();
        configure.customPolicy("servEnvList", policy);
        XWPFTemplate template = XWPFTemplate.compile(templatePath, configure).render(
                new HashMap<String, Object>() {{
                    put("servEnvList", servEnvList);
                }}
        );
        //4.输出文件
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/输出环境配置.docx");
        template.write(out);
        template.close();
        out.close();
    }
    /**
     * poi测试检测用例模板
     */
    @Test
    void testRenderTemplate4() throws IOException {
        //1. 准备数据
        List<ServEnv> servEnvList = Arrays.asList(
                // 服务器1（id=1）的两个软件
                new ServEnv(1, "Windows系统", "64*1T", "Mysql", "V1.0", "NBKJ", "存储数据"),
                new ServEnv(1, "Windows系统", "64*1T", "Redis", "V2.0", "NBKJ", "缓存数据"),
                new ServEnv(1, "Windows系统", "64*1T", "MQ", "V3.0", "NBKJ", "消息队列"),
                new ServEnv(1, "Linux系统", "32*512G", "Tomcat", "V9.0", "Apache", "Web容器"),
                new ServEnv(1, "Linux系统", "32*512G", "Nginx", "V1.18", "开源", "反向代理")
        );
        //2.配置绑定
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        //3.创建表格模板
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy() {
            @Override
            protected void afterloop(XWPFTable table, Object data) {
                // 先调用父类方法完成数据渲染
                super.afterloop(table, data);

                // 根据ID合并单元格
                mergeCellsById(table);
            }

            /**
             * 核心合并方法：根据ID列的值合并相同ID的行
             * 表格结构：
             * - 行0: 主表头
             * - 行1: 子表头
             * - 行2+: 数据行
             * 列结构：
             * - 列0: ID（序号）
             * - 列1: 硬件名称
             * - 列2: 硬件配置
             * - 列3-5: 软件信息（不合并）
             */
            private void mergeCellsById(XWPFTable table) {
                // 数据行起始位置（跳过2行表头）
                final int DATA_START_ROW = 2;
                int totalRows = table.getNumberOfRows();

                // 没有足够数据行时直接返回
                if (totalRows <= DATA_START_ROW) return;

                // 从第一个数据行开始扫描
                int currentRow = DATA_START_ROW;

                while (currentRow < totalRows) {
                    // 获取当前行的ID值
                    String currentId = getCellText(table, currentRow, 0);
                    if (currentId.isEmpty()) {
                        currentRow++;
                        continue;
                    }

                    // 查找相同ID的结束行
                    int endRow = currentRow;
                    for (int nextRow = currentRow + 1; nextRow < totalRows; nextRow++) {
                        String nextId = getCellText(table, nextRow, 0);
                        if (currentId.equals(nextId)) {
                            endRow = nextRow;
                        } else {
                            break;
                        }
                    }

                    // 如果有多行相同ID，合并这三列
                    if (endRow > currentRow) {
                        // 合并ID列（列0）
                        TableTools.mergeCellsVertically(table, 0, currentRow, endRow);
                        // 合并硬件名称列（列1）
                        TableTools.mergeCellsVertically(table, 1, currentRow, endRow);
                        // 合并硬件配置列（列2）
                        TableTools.mergeCellsVertically(table, 2, currentRow, endRow);

                        // 跳转到下一组数据的开始
                        currentRow = endRow + 1;
                    } else {
                        // 单行数据，跳到下一行
                        currentRow++;
                    }
                }
            }

            /**
             * 安全获取单元格文本
             */
            private String getCellText(XWPFTable table, int row, int col) {
                try {
                    XWPFTableCell cell = table.getRow(row).getCell(col);
                    return cell.getText() != null ? cell.getText().trim() : "";
                } catch (Exception e) {
                    return "";
                }
            }
        };
        //模板路径
        String templatePath = "src/main/resources/templateDoc/02检测用例 - 占位符.docx";
        Configure configure = builder.build();
        configure.customPolicy("servEnvList", policy);
        configure.customPolicy("cliEnvList", policy);
        XWPFTemplate template = XWPFTemplate.compile(templatePath, configure).render(
                new HashMap<String, Object>() {{
                    put("servEnvList", servEnvList);
                    put("cliEnvList", servEnvList);
                    put("planNumber", 999999999);
                    put("sampleName", "CODEX");
                    put("sampleVersion", "V1.0");
                    put("juristicUser", "林嘉鹏");
                    put("auditUser", "林嘉鹏");
                    put("approveUser", "林嘉鹏");
                    put("juristicDate", "2025年12月5日");
                    put("auditDate", "2025年12月5日");
                    put("approveDate", "2025年12月5日");
                }}
        );
        //4.输出文件
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/输出02检测用例 - 占位符.docx");
        template.write(out);
        template.close();
        out.close();
    }
    @Test
    void testRenderTemplateWithJson1() throws IOException {
        //1. 读取json数据
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> envConfigList=mapper.readValue(
                new File("src/main/resources/json/envConfig.json"),
                new TypeReference<Map<String, Object>>() {}
        );
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        Configure configure = builder.build();
        configure.customPolicy("envConfigList", policy);
        String path="src/main/resources/templateDoc/02检测用例 - 占位符.docx";
        XWPFTemplate template=XWPFTemplate.compile(path, configure).render(envConfigList);
        template.write(new FileOutputStream("src/main/resources/outputDoc/输出检测用例1.docx"));
        template.close();
    }
    @Test
    void testRenderLoadXReport() throws IOException {
        // 1. 准备被渲染的数据
        Map<String, Object> data = new HashMap<>();
        // 文档基本信息
        data.put("loadXPlanName", "LoadX测试报告-11111111");
        data.put("loadXPlanExcuteDate", "2025年12月8日至2025年12月9日");
        data.put("loadXPlanExcuteDuration", "24小时");
        data.put("loadXConcurrentUsers", "12222人");
        //文档中嵌套的列表
        List<Map<String,String>>loadXRecordList=new ArrayList<>();
        Map<String,String>loadXRecordMap1=new HashMap<>();
        loadXRecordMap1.put("loadRecordId", "1");
        loadXRecordMap1.put("loadXRecordTransaction", "登录");
        loadXRecordMap1.put("loadXRecordUserCount", "1000");
        loadXRecordMap1.put("loadXaverageResponseTime", "1小时");
        loadXRecordMap1.put("loadXthroughput", "1024B");
        loadXRecordMap1.put("loadXRecordsuccessRate", "90%");
        loadXRecordList.add(loadXRecordMap1);
        Map<String,String>loadXRecordMap2=new HashMap<>();
        loadXRecordMap2.put("loadRecordId", "2");
        loadXRecordMap2.put("loadXRecordTransaction", "注册");
        loadXRecordMap2.put("loadXRecordUserCount", "5000");
        loadXRecordMap2.put("loadXaverageResponseTime", "11小时");
        loadXRecordMap2.put("loadXthroughput", "2048B");
        loadXRecordMap2.put("loadXRecordsuccessRate", "80%");
        loadXRecordList.add(loadXRecordMap2);
        Map<String,String>loadXRecordMap3=new HashMap<>();
        loadXRecordMap3.put("loadRecordId", "3");
        loadXRecordMap3.put("loadXRecordTransaction", "搜索");
        loadXRecordMap3.put("loadXRecordUserCount", "2000");
        loadXRecordMap3.put("loadXaverageResponseTime", "3小时");
        loadXRecordMap3.put("loadXthroughput", "1024B");
        loadXRecordMap3.put("loadXRecordsuccessRate", "50%");
        loadXRecordList.add(loadXRecordMap3);
        data.put("loadXRecordList", loadXRecordList);
        // 图片（如果有）
        FileInputStream fis = new FileInputStream("src/main/resources/image/poi测试图片.png");
        PictureRenderData picture = Pictures.ofStream(fis, PictureType.PNG)
                .size(400, 200)
                .create();
        data.put("loadXActiveUserCount", picture);
        data.put("loadXRequestCountPerSecond", picture);
        data.put("loadXTransactionCountPerSecond", picture);
        data.put("loadXThroughputPerSecond", picture);
        data.put("loadXAverageResponseTime", picture);
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        Configure config = Configure.builder()
                .buildGramer("${", "}")  // 使用 ${} 作为标签前后缀
                .build();
        config.customPolicy("loadXRecordList", policy);
        //准备模板
        XWPFTemplate template = XWPFTemplate
                .compile("src/main/resources/templateDoc/LoadX测试报告.docx", config)
                .render(data);
        //输出文档
        template.writeAndClose(new FileOutputStream("src/main/resources/outputDoc/LoadX测试报告输出2.docx"));
        System.out.println("文档生成成功！");
    }

}
