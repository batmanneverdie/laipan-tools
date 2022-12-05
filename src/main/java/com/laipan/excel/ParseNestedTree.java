package com.laipan.excel;

import com.laipan.entity.Result;
import com.laipan.factory.ExcelFactory;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import java.util.Optional;

import static com.laipan.utils.ExcelUtils.getCellValueByCell;

@RestController
@RequestMapping("/tools/excel")
public class ParseNestedTree {

    private static int next_node_lft = 1;
    private static int additionInfoStartCol;
    private static int additionInfoEndCol;

    private static int nodeWidth;
    private static int treeDepth;
    private static int startRow;
    private static int endRow;

    private static Sheet sheet;

    private static Integer nodeNum; // 节点个数

    @PostMapping("generateTDT")
    public Result<Object> getMemberList(MultipartHttpServletRequest request) {
        // 1. 获取 excel sheet、起止行
        final int sheetIndex = Optional
                .ofNullable(request.getParameter("sheetIndex"))
                .map(str -> Integer.parseInt(request.getParameter("sheetIndex")) - 1).orElse(0);
        startRow = Integer.parseInt(request.getParameter("startRow")) - 1;
        endRow = Integer.parseInt(request.getParameter("endRow")) - 1;
        nodeWidth = Integer.parseInt(request.getParameter("nodeWidth"));
        treeDepth = Integer.parseInt(request.getParameter("treeDepth"));
        final int additionInfoWidth = Integer.parseInt(request.getParameter("additionInfoWidth"));

        // 1.1 计算 node 节点的宽度
        additionInfoStartCol = nodeWidth * treeDepth;
        additionInfoEndCol = nodeWidth * treeDepth + additionInfoWidth;

        // 2. 获取表单中文件数据
        System.out.println("---------获取表单中文件数据---------");
        MultiValueMap<String, MultipartFile> multiFileMap = request.getMultiFileMap();

        // 3. 遍历表单中元素信息
        List<MultipartFile> upfiles = multiFileMap.get("file");
        MultipartFile       file    = upfiles.get(0);

        try {
            // 4. 解析 excel
            sheet = ExcelFactory.createSheet(file.getOriginalFilename(), file.getInputStream(), sheetIndex);

            // 获取树的最上根节点（单颗树）
            TreeNode parentNode = new TreeNode();
            parentNode.setRowNo(startRow);
            parentNode.setSameDepthNextNodeRowNo(endRow);
            recursiveGetTreeNodeByDepth(parentNode, 1, endRow, 0);

            // 遍历根节点，获取树结构
            for (TreeNode parentTreeNode : parentNode.getChildren()) {
                // 递归获取子节点，建立树结构
                recursiveGetTreeNodeByDepth(parentTreeNode, parentTreeNode.getDepth() + 1, parentTreeNode.getSameDepthNextNodeRowNo(), treeDepth);
                // 递归获取每个节点的所有子节点个数
                recursiveGetChildrenNum(parentTreeNode);
                // 递归生成 lft、rgt 编号并获取节点的附加信息
                recursion(parentTreeNode);
                // 父节点结束应该 + 1
                next_node_lft = parentTreeNode.getRgt() + 1;
            }

            // 重置 next_node_lft
            next_node_lft = 1;

            // 5. 输出 csv, 生成二维表
            System.out.println(parentNode.getChildren());
        } catch (IOException e) {
            e.printStackTrace();
            return Result.error(1, "解析失败！");
        }
        return Result.OK();
    }


    /**
     * 递归获取子节点，同时获取每个节点的高度（excel 的行数）
     *
     * @param rootNode   根节点
     * @param depth      层级
     * @param treeEndRow 根节点高度
     */
    public static void recursiveGetTreeNodeByDepth(TreeNode rootNode, int depth, int treeEndRow, int recursionTimes) {
        // 获取根节点的下级子节点
        LinkedList<TreeNode> childNodes = new LinkedList<>();
        // 遍历每一行的数据
        int treeNodeIndex = 0;
        for (int rowNo = rootNode.getRowNo(); rowNo <= rootNode.getSameDepthNextNodeRowNo(); rowNo++) {
            Row row = sheet.getRow(rowNo);
            // 遍历每一行此层级的数据
            if (getCellValueByCell(row.getCell((depth - 1) * nodeWidth)) != null) {
                TreeNode treeNode = new TreeNode();
                treeNode.setDepth(depth);
                treeNode.setStartCol((depth - 1) * nodeWidth);
                treeNode.setRowNo(rowNo);
                treeNode.setSameDepthNextNodeRowNo(treeEndRow);

                int loopCount = 0;
                for (int colNo = (depth - 1) * nodeWidth; colNo <= nodeWidth * depth; colNo++) {
                    if (loopCount == 1) {
                        treeNode.setOrd(Integer.parseInt(getCellValueByCell(row.getCell(colNo))));
                    }
                    if (loopCount == 2) {
                        treeNode.setExpenseType(getCellValueByCell(row.getCell(colNo)));
                    }
                    loopCount++;
                }
                childNodes.add(treeNode);

                // 获取树的高度（excel 的行数）
                if (treeNodeIndex != 0) {
                    childNodes.get(treeNodeIndex - 1).setSameDepthNextNodeRowNo(rowNo - 1);
                }
                treeNodeIndex++;
            }
        }
        rootNode.setChildren(childNodes);

        // 控制递归的入口
        if (depth < recursionTimes) {
            for (TreeNode childNode : childNodes) {
                recursiveGetTreeNodeByDepth(childNode, depth + 1, childNode.getSameDepthNextNodeRowNo(), recursionTimes);
            }
        }
    }

    /**
     * 递归获取根节点下所有子节点的个数
     */
    public static void recursiveGetChildrenNum(TreeNode rootNode) {
        nodeNum = 0;
        // todo 此处的 if 语句是否可以省略
        if (CollectionUtils.isNotEmpty(rootNode.getChildren())) {
            for (TreeNode child : rootNode.getChildren()) {
                recursiveChildren(rootNode);
                rootNode.setChildNodeNum(nodeNum);
                recursiveGetChildrenNum(child);
            }
        }
    }

    /**
     * 递归获取根节点下所有子节点
     */
    public static void recursiveChildren(TreeNode rootNode) {
        // todo 此处的 if 语句是否可以省略
        if (CollectionUtils.isNotEmpty(rootNode.getChildren())) {
            for (TreeNode child : rootNode.getChildren()) {
                nodeNum++;
                // 递归所有节点，与当前节点比较是否在当前节点的子节点
                recursiveChildren(child);
            }
        }
    }

    /**
     * 递归，根左右遍历生成 lft、rgt 编号，获取附加信息
     *
     * @param root 节点
     */
    public static void recursion(TreeNode root) {
        root.setLft(next_node_lft);
        root.setRgt(root.getLft() + root.getChildNodeNum() * 2 + 1);

        // 获取附加信息
        if (root.getLft() + 1 == root.getRgt()) {
            // 只有叶子节点才会挂附加信息
            LinkedList<NodeRecord> nodeRecords = new LinkedList<>();
            for (int additionalRowNo = root.getRowNo(); additionalRowNo <= root.getSameDepthNextNodeRowNo(); additionalRowNo++) {
                Row row = sheet.getRow(additionalRowNo);

                // 叶子节点是整颗树的根节点时，直接判断有无附加信息
                if (root.getStartCol() != 0) {
                    // 判断同一层级的下一行的左边是否有值，如果有则终止
                    if (additionalRowNo != root.getRowNo())
                        if (getCellValueByCell(row.getCell(root.getStartCol() - 2)) != null) break;
                }

                if (getCellValueByCell(row.getCell(additionInfoStartCol)) == null) break;
                NodeRecord nodeRecord = new NodeRecord();
                int        loopCount  = 1;
                for (int additionalColNo = additionInfoStartCol; additionalColNo <= additionInfoEndCol; additionalColNo++) {
                    if (loopCount == 1) {
                        nodeRecord.setCourseCode(getCellValueByCell(row.getCell(additionalColNo)));
                    }
                    if (loopCount == 2) {
                        nodeRecord.setCourseName(getCellValueByCell(row.getCell(additionalColNo)));
                    }
                    if (loopCount == 3) {
                        nodeRecord.setExamineFlag(getCellValueByCell(row.getCell(additionalColNo)));
                    }
                    if (loopCount == 4) {
                        nodeRecord.setComment(getCellValueByCell(row.getCell(additionalColNo)));
                    }
                    if (loopCount == 5) {
                        nodeRecord.setCourseType(getCellValueByCell(row.getCell(additionalColNo)));
                    }
                    loopCount++;
                }

                nodeRecords.add(nodeRecord);
            }
            root.setRecords(nodeRecords);
        }

        // 打印节点信息
//        for (int i = 1; i < root.getDepth(); i++) {
//            System.out.print("\t");
//        }
//        System.out.printf("depth: %d, \tlft_no: %d, \trgt_no: %d, \t%d--%s.\n", root.getDepth(), root.getLft(), root.getRgt(), root.getOrd(), root.getExpenseType());
        if (CollectionUtils.isEmpty(root.getRecords())) {
            System.out.printf("%d\t%d\t%s\t%d\n", root.getLft(), root.getOrd(), root.getExpenseType(), root.getRgt());
        } else {
            for (NodeRecord record : root.getRecords()) {
                System.out.printf("%d\t%d\t%s\t%d\t", root.getLft(), root.getOrd(), root.getExpenseType(), root.getRgt());
                System.out.printf("%s\t%s\t%s\t%s\t%s\n", record.getCourseCode(), record.getCourseName(), record.getExamineFlag(), record.getComment(), record.getExamineFlag());
            }
        }

        // 递归子节点
        // TODO if 判断
        if (CollectionUtils.isNotEmpty(root.getChildren())) {
            for (TreeNode treeNode : root.getChildren()) {
                // 判断有无子节点，有 + 1，无 + 2
                if (root.getChildren().size() == 0) {
                    next_node_lft += 2;
                } else {
                    next_node_lft++;
                }
                recursion(treeNode);
            }
        }
        // 每一层级的循环结束，lft 应该 + 1
        next_node_lft++;
    }
}
