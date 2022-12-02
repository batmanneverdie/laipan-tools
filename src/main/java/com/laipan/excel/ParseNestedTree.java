package com.laipan.excel;

import com.laipan.entity.Result;
import com.laipan.factory.ExcelFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;

@RestController
@RequestMapping("/tools/excel")
public class ParseNestedTree {

    @PostMapping("generateTDT")
    public Result<Object> getMemberList(MultipartHttpServletRequest request) {
        // 1. 获取 excel sheet、起止行
        final int sheetIndex = Optional
                .ofNullable(request.getParameter("sheetIndex"))
                .map(str -> Integer.parseInt(request.getParameter("sheetIndex")) - 1).orElse(0);
        final int startRow          = Integer.parseInt(request.getParameter("startRow")) - 1;
        final int endRow            = Integer.parseInt(request.getParameter("endRow")) - 1;
        final int nodeWidth         = Integer.parseInt(request.getParameter("nodeWidth"));
        final int treeDepth         = Integer.parseInt(request.getParameter("treeDepth"));
        final int additionInfoWidth = Integer.parseInt(request.getParameter("additionInfoWidth"));

        // 1.1 计算 node 节点的宽度
        final int allNodeWith          = nodeWidth * treeDepth;
        int       nodeStartColNo       = 0;
        final int nodeEndColNo         = allNodeWith;
        final int additionInfoStartCol = allNodeWith;
        final int additionInfoEndCol   = allNodeWith + additionInfoWidth;

        // 2. 获取表单中文件数据
        System.out.println("---------获取表单中文件数据---------");
        MultiValueMap<String, MultipartFile> multiFileMap = request.getMultiFileMap();

        // 3. 遍历表单中元素信息
        List<MultipartFile> upfiles = multiFileMap.get("file");
        MultipartFile       file    = upfiles.get(0);

        ArrayList<TreeNode> treeNodes = new ArrayList<>();
        try {
            // 4. 解析 excel
            Sheet sheet = ExcelFactory.createSheet(file.getOriginalFilename(), file.getInputStream(), sheetIndex);

            // 4.1 遍历每一层级的节点
            for (int depth = 1; depth <= treeDepth; depth++) {
                // 4.1.1 遍历每一行的数据
                for (int rowNo = startRow; rowNo <= endRow; rowNo++) {
                    Row row = sheet.getRow(rowNo);

                    // 4.1.1.1 遍历每一行此时层级的数据
                    if (getCellValueByCell(row.getCell(nodeStartColNo)) != null) {
                        TreeNode treeNode = new TreeNode();
                        treeNode.setRowNo(rowNo);
                        treeNode.setStartCol(nodeStartColNo);

                        ArrayList<NodeRecord> nodeRecords = new ArrayList<>();
                        NodeRecord nodeRecord = new NodeRecord();
                        nodeRecords.add(nodeRecord);

                        treeNode.setRecords(nodeRecords);
                        int loopCount = 0;
                        for (int colNo = nodeStartColNo; colNo <= nodeWidth * depth; colNo++) {
                            if (loopCount == 1) {
                                nodeRecord.setOrd(Integer.parseInt(getCellValueByCell(row.getCell(colNo))));
                            }
                            if (loopCount == 2) {
                                nodeRecord.setExpenseType(getCellValueByCell(row.getCell(colNo)));
                            }
                            loopCount++;
                        }
                        loopCount = 0;
                        treeNodes.add(treeNode);
                    }
                }
                nodeStartColNo = nodeWidth * depth;
            }
            treeNodes.forEach(System.out::println);



            //                    // 4.1.1.1.1 判断此节点是否直接挂附加信息列
//                    boolean existAdditionalInfo = false;
//                    int additionalInfoRowNum = 0;
//                    for (int colNo2 = nodeStartColNo; colNo2 <= additionInfoStartCol; colNo2++) {
//                        if (getCellValueByCell(row.getCell(colNo2)) == null) {
//                            existAdditionalInfo = true;
//                        }
//                        // 一个费用类型下可挂多个科目，需要判断下面几行的科目是否也属于此费用类型
//                        additionalInfoRowNum++;
//                        break;
//                    }
//                    if (existAdditionalInfo) {
//                        // 存在附加信息
//                        System.out.print("exist additional info: ");
//                        // 获取附加信息
//                        for (int colNo3 = additionInfoStartCol; colNo3 <= additionInfoEndCol; colNo3++) {
//                            System.out.print(getCellValueByCell(row.getCell(colNo3)) + "---");
//                        }
//                        System.out.println();
//                    }

            // 5. 生成二维表

        } catch (IOException e) {
            e.printStackTrace();
            return Result.error(1, "解析失败！");
        }
        return Result.OK();
    }

    //获取单元格各类型值，返回字符串类型
    public static String getCellValueByCell(Cell cell) {
        //判断是否为 null 或空串，空串也当作 null
        if (cell == null || cell.toString().trim().equals("")) {
            return null;
        }
        String   cellValue = "";
        CellType cellType  = cell.getCellType();
        switch (cellType) {
            case NUMERIC: // 数字
                short format = cell.getCellStyle().getDataFormat();
                if (DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = null;
                    if (format == 20 || format == 32) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else if (format == 14 || format == 31 || format == 57 || format == 58) {
                        // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = org.apache.poi.ss.usermodel.DateUtil
                                .getJavaDate(value);
                        cellValue = sdf.format(date);
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }
                    try {
                        cellValue = sdf.format(cell.getDateCellValue());// 日期
                    } catch (Exception e) {
                        try {
                            throw new Exception("exception on get date data !".concat(e.toString()));
                        } catch (Exception e1) {
                            e1.printStackTrace();
                        }
                    } finally {
                        sdf = null;
                    }
                } else {
                    BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
                    cellValue = bd.toPlainString();// 数值 这种用BigDecimal包装再获取plainString，可以防止获取到科学计数值
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                // 如果直接获取公式类型的数据，则获取直接就是公式本身，如：1+1，""员工"" 等
                if (cell.getCellFormula().startsWith("EOMONTH")) {
                    // Excel 的 EOMONTH 函数在计算日期不会返回 string 类型的日期数据，而是返回日期序列号，1900-01-01 的序列号为 1，按日递增。
                    SimpleDateFormat sdf = null;
                    try {
                        Date date = cell.getDateCellValue();
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = sdf.format(date);
                    } catch (Exception e) {
                        try {
                            throw new Exception("exception on get date data within 【EOMONTH】 function !".concat(e.toString()));
                        } catch (Exception e1) {
                            e1.printStackTrace();
                        }
                    } finally {
                        sdf = null;
                    }
                } else {
                    cellValue = ((XSSFCell) cell).getRawValue();
                }
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                cellValue = "ERROR VALUE";
                break;
            default:
                cellValue = "UNKNOW VALUE";
                break;
        }
        return cellValue;
    }
}
