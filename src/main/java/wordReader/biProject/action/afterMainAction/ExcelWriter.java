package wordReader.biProject.action.afterMainAction;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import wordReader.biProject.excelFormat.CusCell;
import wordReader.biProject.excelFormat.CusCellStyle;
import wordReader.biProject.model.DataPojo;
import wordReader.biProject.model.PinkPojo;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelWriter {

    private static final List<String> CELL_HEADS; //列頭
    private static final List<String> PINK_HEADS; //列頭
    private static final List<String> BELOW_HEADS;

    // 欄位名稱
    static {
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("編號");
        CELL_HEADS.add("部門");
        CELL_HEADS.add("員工姓名");
        CELL_HEADS.add("實際申請時間");
        CELL_HEADS.add("(起)");
        CELL_HEADS.add("實際申請時間");
        CELL_HEADS.add("(迄)");
        CELL_HEADS.add("申請日期");
        CELL_HEADS.add("申請時數");
        CELL_HEADS.add("專案");
        CELL_HEADS.add("事由");
        CELL_HEADS.add("承認工時");
        CELL_HEADS.add("使用方式");
        CELL_HEADS.add("實際上班時間");
        CELL_HEADS.add("實際下班時間");
        CELL_HEADS.add("對照申請時數");
        CELL_HEADS.add("缺卡內容");
        CELL_HEADS.add("備註");

        PINK_HEADS = new ArrayList<>();
        PINK_HEADS.add("編號");
        PINK_HEADS.add("日期");
        PINK_HEADS.add("員工");
        PINK_HEADS.add("上班");
        PINK_HEADS.add("下班");
        PINK_HEADS.add("缺卡內容");

        BELOW_HEADS = new ArrayList<>();
        BELOW_HEADS.add("編號");
        BELOW_HEADS.add("部門");
        BELOW_HEADS.add("員工姓名");
        BELOW_HEADS.add("日期");
        BELOW_HEADS.add("專案");
        BELOW_HEADS.add("承認工時");
        BELOW_HEADS.add("使用方式");
        BELOW_HEADS.add("加班一階\n個數\n(1-2H)");
        BELOW_HEADS.add("加班二階\n個數\n(3-8H)");
        BELOW_HEADS.add("加班三階\n個數\n(9-12H)");
        BELOW_HEADS.add("加班補休");
        BELOW_HEADS.add("補休滿6H給1天；國定假日滿1H給1天；\n"
                + "超過8H依規定計算；\n" + "例假日滿1H給1天+補假1天(匯入獎勵休假)\n"
                + "若加班為3-4H，二階個數為2H\n" + "若加班為6-8H天，二階個數為6H");
        BELOW_HEADS.add("承認天數");
    }

    /**
     * 生成Excel並寫入數據信息 ,這個function等於是ExcelWriter 的 Main Class
     *
     * @param dataList  數據列表
     * @param pinkPojos excel數據列表
     * @return 寫入數據後的工作簿對象
     */
    @SuppressWarnings("unused")
    public static Workbook exportData(List<DataPojo> dataList, List<PinkPojo> pinkPojos) {
        // 生成xlsx的Excel
        Workbook workbook = new SXSSFWorkbook();
        // 如需生成xls的Excel，請使用下面的工作簿對象，注意後續輸出時文件後綴名也需更改為xls
        // Workbook workbook = new HSSFWorkbook();

        // 生成Sheet表，並且建立表頭
        Sheet sheet = buildDataSheet(workbook);

        // 填上第一個主表內容
        int rowNum = 5; // 因為表頭是建立在index 4的地方
        for (DataPojo data : dataList) {
            //輸出行數據
            Row row = sheet.createRow(rowNum++);
            // 工作單表格
            convertDataToRow(workbook, sheet, data, row);
        }
        // 合併“編號儲存格
        handleIdn(sheet);


        // 在主table底下再建立一個table
        rowNum = rowNum + 3;
        secondTableHeader(sheet, rowNum++);
        int startOfTable2BodyRowNum = rowNum;
        for (DataPojo data : dataList) {
            Row row = sheet.createRow(rowNum++);

            convertDataToSecondTableRow(workbook, sheet, data, row);
        }

        // 合併編號
        handleTable2Idn(sheet, startOfTable2BodyRowNum, rowNum);
        // 合併使用方式 及承認天數
        handleTable2UsewaysNAdmitDays(sheet, startOfTable2BodyRowNum, rowNum);

        // 加班明細sheet
        Sheet sheet2 = buildDataSheet2(workbook, pinkPojos);

        return workbook;
    }

    /**
     * 生成sheet表1，並寫入第一行數據（列頭）and 相關說明 更新日期等
     *
     * @param workbook 工作簿對象
     * @return 已經寫入列頭的Sheet
     */
    private static Sheet buildDataSheet(Workbook workbook) {
        // createSheet裡面放參數 可指定工作表名稱
        Sheet sheet = workbook.createSheet("加班明細表");
        // 設置 Column 寬度、高度
        setColumnOfSheet(sheet);

        // 標題 : 月份員工加班明細表
        // 取得範圍內的儲存格 (start row , end row , start column , end column )
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 1, 12);
        // 合併儲存格
        sheet.addMergedRegion(cellRangeAddress);
        // 創建標題row
        Row title = sheet.createRow(0);
        title.setHeight((short) 800); // title row height : 800
        Cell titleCell = title.createCell(1);
        titleCell.setCellValue("月份員工加班明細表"); // 標題
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 30); // 字體大小設定為30
        font.setBold(true); // 粗體
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        titleStyle.setFont(font); // 塞入字體風格
        titleCell.setCellStyle(titleStyle); // 塞入style到cell

        // 標題右邊說明
        cellRangeAddress = new CellRangeAddress(0, 2, 13, 15);
        sheet.addMergedRegion(cellRangeAddress);
        Cell detailCell = title.createCell(13);
        detailCell.setCellValue("109年確認，承認工時需扣除12-13、18-19吃飯時間(半小時不作扣除)，配合以下勞動基準法 - 勞動條件及就業平等目：\n" +
                "第 35 條-勞工繼續工作四小時，至少應有三十分鐘之休息。但實行輪班制或其工作有連續性或緊急性者，雇主得在工作時間內，另行調配其休息時間。");
        Font fontSizeTenFont = workbook.createFont();
        fontSizeTenFont.setFontHeightInPoints((short) 13);
        CellStyle detCellStyle = workbook.createCellStyle();
        detCellStyle.setWrapText(true);
        detCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        detCellStyle.setFont(fontSizeTenFont);
        detailCell.setCellStyle(detCellStyle);

        // 標題下面: 更新日期
        cellRangeAddress = new CellRangeAddress(2, 2, 1, 8);
        sheet.addMergedRegion(cellRangeAddress);
        Row updateTimeRow = sheet.createRow(2);
        updateTimeRow.setHeight((short) 500);
        Cell updateTimeCellString = updateTimeRow.createCell(1);
        updateTimeCellString.setCellValue("更新日期:"); // 內容
        CellStyle updateTimeCellStringStyle = workbook.createCellStyle();
        updateTimeCellStringStyle.setAlignment(HorizontalAlignment.RIGHT); // 靠右
        updateTimeCellStringStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        updateTimeCellStringStyle.setFont(fontSizeTenFont);
        updateTimeCellString.setCellStyle(updateTimeCellStringStyle);

        // 更新日期
        cellRangeAddress = new CellRangeAddress(2, 2, 9, 10);
        sheet.addMergedRegion(cellRangeAddress);
        Cell updateTimeCell = updateTimeRow.createCell(9);
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");
        LocalDateTime now = LocalDateTime.now();
        updateTimeCell.setCellValue(dtf.format(now)); // 時間
        CellStyle updateTimeStyle = workbook.createCellStyle();
        updateTimeStyle.setAlignment(HorizontalAlignment.LEFT); // 靠左
        updateTimeStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        updateTimeStyle.setFont(fontSizeTenFont);
        updateTimeCell.setCellStyle(updateTimeStyle);


        CellRangeAddress cellRangeAddress1 = new CellRangeAddress(4, 4, 9, 10);
        sheet.addMergedRegion(cellRangeAddress1);

        CellRangeAddress cellRangeAddress2 = new CellRangeAddress(4, 4, 11, 13);
        sheet.addMergedRegion(cellRangeAddress2);

        // 構建主表頭樣式
        CellStyle cellStyle = CusCellStyle.buildHeadCellStyle(sheet.getWorkbook());
        // 寫入第一行各列的數據
        Row head = sheet.createRow(4);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell;
            if (i == 10) {
                cell = head.createCell(i + 1); // 11
            } else if (i > 10) {
                cell = head.createCell(i + 3); // 11
            } else {
                cell = head.createCell(i);
            }

            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }

        // 補上合併儲存格的邊筐缺陷
        setCombineRegionBorder(sheet, cellRangeAddress1);
        setCombineRegionBorder(sheet, cellRangeAddress2);

        return sheet;
    }

    /**
     * 設置Column高度，依照表頭內容設定寬度
     *
     * @param sheet
     */
    private static void setColumnOfSheet(Sheet sheet) {
        int offset = 0; // 因為有些cell有合併，因此需要offSet來控制Index
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            switch (CELL_HEADS.get(i)) {
                case "專案":
                    sheet.setColumnWidth(i, 3500);
                    offset = offset + 1;
                    sheet.setColumnWidth(i + offset, 3500);
                    break;
                case "事由":
                    sheet.setColumnWidth(i + offset, 4000);
                    offset = offset + 1;
                    sheet.setColumnWidth(i + offset, 4000);
                    offset = offset + 1;
                    sheet.setColumnWidth(i + offset, 13000);
                    break;
                case "(起)":
                case "(迄)":
                    sheet.setColumnWidth(i, 2500);
                    break;
                case "實際上班時間":
                case "實際下班時間":
                case "對照申請時數":
                case "缺卡內容":
                    sheet.setColumnWidth(i + offset, 7000);
                    break;
                case "承認工時":
                case "使用方式":
                    sheet.setColumnWidth(i + offset, 5000);
                    break;
                case "備註":
                    sheet.setColumnWidth(i + offset, 15000);
                    break;
                case "部門":
                    sheet.setColumnWidth(i, 10000);
                    break;
                default:
                    sheet.setColumnWidth(i, 5000);
                    break;
            }
        }

        // 設置默認行高
        sheet.setDefaultRowHeight((short) 400);
    }

    /**
     * 將數據轉換成行
     *
     * @param data 源數據
     * @param row  行對象
     * @return
     */
    private static void convertDataToRow(Workbook workbook, Sheet sheet, DataPojo data, Row row) {
        int cellNum = 0;
        int offset = 0;
        Cell cell;
        boolean yellowBack = false;
        if (data.isSunday()) yellowBack = true;

        // 編號
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
        cell.setCellValue(row.getRowNum()); // error ?

        // 部門
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getApartment() == null ? "" : data.getApartment());

        // 員工姓名
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getName() == null ? "" : data.getName());

        // 實際申請時間
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getStartDay() == null ? "" : data.getStartDay());

        // (起)
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getStartTime() == null ? "" : data.getStartTime());

        // 實際申請時間
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getEndDay() == null ? "" : data.getEndDay());

        // (迄)
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getEndTime() == null ? "" : data.getEndTime());

        // 申請日期
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getDate() == null ? "" : data.getDate());

        // 申請時數
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getApplyHour() == null ? "" : data.getApplyHour());

        // 專案
        CellRangeAddress cra = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 9, 10);
        sheet.addMergedRegion(cra);
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getProjectName() == null ? "" : data.getProjectName());
        offset = offset + 1;

        // 事由
        cra = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 11, 13);
        sheet.addMergedRegion(cra);
        cell = CusCell.createCellWithLeftAlignment(workbook, row, cellNum + offset, yellowBack);
        cell.setCellValue(data.getReason() == null ? "" : data.getReason());
        cellNum = cellNum + 1;
        offset = offset + 2;

        // 承認工時
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
//	   	redFontCellStyle
        if (data.getAdmitTime() != null) {
            if (redRule(data.getAdmitTime())) {
                // 要標紅色
                CellStyle redFontCellStyle = CusCellStyle.redFontCellStyle(workbook);
                cell.setCellStyle(redFontCellStyle);
            }

            cell.setCellValue(data.getAdmitTime());
        } else
            cell.setCellValue("");

        cellNum = cellNum + 1;

        // 使用方式
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getRestOrMoney());
        cellNum = cellNum + 1;

        // 實際上班時間
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getActualStartTime());
        cellNum = cellNum + 1;

        // 實際下班時間
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getActualEndTime());
        cellNum = cellNum + 1;

        // 對照申請時數
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        String hourString = String.valueOf((data.getDifferTotalTime() / 60));
        String minString = String.valueOf(data.getDifferTotalTime() % 60);
        cell.setCellValue(hourString + "時" + minString + "分");
        cellNum = cellNum + 1;

        // 缺卡內容
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getMissContent());
        cellNum = cellNum + 1;

        // 備註
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getExtraMsg());
//	   	 cellNum = cellNum + 1 ;

        CellRangeAddress cellRangeAddress1 = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 9, 10);
        CellRangeAddress cellRangeAddress2 = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 11, 13);
        // 補上合併儲存格的邊筐缺陷
        setCombineRegionBorder(sheet, cellRangeAddress1);
        setCombineRegionBorder(sheet, cellRangeAddress2);

    }

    /**
     * 處理前面編號合併部分, 名稱部分相同的row 編號column就會合併
     *
     * @param sheet
     */
    private static void handleIdn(Sheet sheet) {

        int columnIndex = 0; // for 尋找A欄位裡面, “編號”是從第幾格開始
        int rowCountLast = sheet.getLastRowNum() + 1; // rowCount 的最大值
        int rowCountFirst = 0; // rowCount 開始值, “編號”下面

        try {
            // loop through A column and find cell with value : 編號
            for (int rowIndex = 0; rowIndex < rowCountLast; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        if (cell.getStringCellValue().equals("編號")) {
//						    	  System.out.println("編號 在第" + (rowIndex+1) + "行" ) ;
                            rowCountFirst = rowIndex + 1;
                        }
                    }
                }
            }
        } catch (Exception e) {
//				  System.out.println("Cannot get a STRING value from a NUMERIC cell 可忽略這個在尋找編號欄位的錯誤.");
        }

        // 在接到第一次error之後就會跳出for迴圈
        // now, 從rowCountFirst到rowCountLast, 找出Column C裡面名字相同的row並把A欄位合併
        columnIndex = 2; // 尋找Column C
        try {
            int idn = 1; // 編號初始號碼
            boolean getNameTime = true;

            Cell baseCell;
            Cell baseIdnCell = null;
            String baseName = "";
            int baseIndex = 0;

            Cell nextCell;
            String nextName;

            for (int i = rowCountFirst; i < rowCountLast; i++) {


                if (getNameTime) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        baseCell = row.getCell(columnIndex);
                        if (baseCell != null) {
                            baseIdnCell = row.getCell(0);

                            baseName = baseCell.getStringCellValue(); // 先找到第一個名子
                            baseIndex = i; // 紀錄Index
                            getNameTime = false; // 關閉找基底名字開關
                        }
                    }
                }

                if (!getNameTime) {
                    if ((i + 1) < rowCountLast) {
                        Row row = sheet.getRow(i + 1);
                        if (row != null) {
                            nextCell = row.getCell(columnIndex);
                            if (nextCell != null) {
                                nextName = nextCell.getStringCellValue(); // 取得名字
                                if (!nextName.equals(baseName)) { // 名字不相同


                                    if (i == baseIndex)
                                        // 不用合併 因為只有一格
                                        // 設定編號號碼即可
                                        baseIdnCell.setCellValue(idn++);
                                    else {
                                        // 代表 要做合併了 從baseIndex到nextIndex
                                        CellRangeAddress cra = new CellRangeAddress(baseIndex, i, 0, 0);
                                        sheet.addMergedRegion(cra);
                                        // 設定編號號碼, 都設定到baseIndex那個cell去
                                        baseIdnCell.setCellValue(idn++);
                                    }

                                    //把getNameTime重新開啟
                                    getNameTime = true;
                                }
                            }
                        }
                    } // end if()
                    else {
                        if (baseIndex == i) baseIdnCell.setCellValue(idn);
                        else {
                            try {
                                CellRangeAddress cra = new CellRangeAddress(baseIndex, i, 0, 0);
                                sheet.addMergedRegion(cra);
                            } catch (Exception e) {
                                System.out.println(e.getMessage());
                                System.out.println("在合併儲存格,範圍沒有包含2個以上的cells導致,當加班筆數少於一定數量會有這個例外,如果結果正確可以忽略");
                            }

                            baseIdnCell.setCellValue(idn);
                        }
                    }
                }

            }

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }

    /**
     * 合併儲存格邊筐設定補足
     *
     * @param sheet
     * @param cra
     */
    private static void setCombineRegionBorder(Sheet sheet, CellRangeAddress cra) {
        RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
    }


    // Sheet1 Table2

    /**
     * 主表下面第二個Table Header生成
     *
     * @param sheet
     * @param rowNum
     */
    private static void secondTableHeader(Sheet sheet, int rowNum) {

        CellRangeAddress cellRangeAddress1 = new CellRangeAddress(rowNum, rowNum, 3, 4);
        sheet.addMergedRegion(cellRangeAddress1);

        CellRangeAddress cellRangeAddress2 = new CellRangeAddress(rowNum, rowNum, 5, 6);
        sheet.addMergedRegion(cellRangeAddress2);

        // 構建頭單元格樣式
        CellStyle cellStyle = CusCellStyle.buildHeadCellStyle(sheet.getWorkbook());

        CellStyle secondCellStyle = CusCellStyle.secondHeadStyle(sheet.getWorkbook());

        CellStyle thirdCellStyle = CusCellStyle.thirdHeadStyle(sheet.getWorkbook());

        CellStyle forthCellStyle = CusCellStyle.fourthHeadStyle(sheet.getWorkbook());
        // 寫入第一行各列的數據
        Row head = sheet.createRow(rowNum);
        head.setHeight((short) 1500); // title row height : 800
        for (int i = 0; i < BELOW_HEADS.size(); i++) {
            Cell cell;

            if (i == 4) {
                cell = head.createCell(i + 1); // 11
            } else if (i >= 5) {
                cell = head.createCell(i + 2); // 11
            } else
                cell = head.createCell(i); // 11

            if (i < 7)
                cell.setCellStyle(cellStyle);
            else if (i <= 10) {
                cell.setCellStyle(secondCellStyle);
            } else if (i == 11) {
                // 51 153 102
                cell.setCellStyle(forthCellStyle);
            } else
                cell.setCellStyle(thirdCellStyle);


            cell.setCellValue(BELOW_HEADS.get(i));
        }

        // 補上合併儲存格的邊筐缺陷
        setCombineRegionBorder(sheet, cellRangeAddress1);
        setCombineRegionBorder(sheet, cellRangeAddress2);

    }

    /**
     * 主表下面第二個Table 內容
     *
     * @param workbook
     * @param sheet
     * @param data
     * @param row
     */
    private static void convertDataToSecondTableRow(Workbook workbook, Sheet sheet, DataPojo data, Row row) {
        int cellNum = 0;
        int offset = 0;
        Cell cell;
        boolean yellowBack = false;
        if (data.isSunday()) yellowBack = true;

        // 編號
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
        cell.setCellValue(row.getRowNum());

        // 部門
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getApartment() == null ? "" : data.getApartment());

        // 員工姓名
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getName() == null ? "" : data.getName());

        // 日期
        CellRangeAddress cra = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 4);
        sheet.addMergedRegion(cra);
        setCombineRegionBorder(sheet, cra);
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, yellowBack);
        cell.setCellValue(data.getStartDay() == null ? "" : data.getStartDay());
        offset++;

        // 專案
        cra = new CellRangeAddress(row.getRowNum(), row.getRowNum(), 5, 6);
        sheet.addMergedRegion(cra);
        setCombineRegionBorder(sheet, cra);
        cell = CusCell.createCellWithLeftAlignment(workbook, row, cellNum + offset, yellowBack);
        cell.setCellValue(data.getProjectName() == null ? "" : data.getProjectName());
        cellNum = cellNum + 1;
        offset++;

        // 承認工時
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, yellowBack);
//		   	  < 6  or > 8 顯示紅色
        if (data.getAdmitTime() != null) {
            if (Integer.parseInt(data.getAdmitTime()) > 8 || Integer.parseInt(data.getAdmitTime()) < 6) {
                // 要標紅色
                CellStyle redFontCellStyle = CusCellStyle.redFontCellStyle(workbook);
                cell.setCellStyle(redFontCellStyle);
            }

            cell.setCellValue(data.getAdmitTime());
        } else
            cell.setCellValue("");

        cellNum = cellNum + 1;

        // 使用方式
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue(data.getRestOrMoney());
        cellNum = cellNum + 1;

        // 承認工時1-5 , 6 - 8 直接算一天8小時, 9-12
        int[] resultHours = returnTotalTime(data.getAdmitTime() == null ? 0 : Integer.parseInt(data.getAdmitTime()));
        if (data.getRestOrMoney().equals("補休")) {
            for (int i = 0; i < 3; i++) {
                cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, yellowBack);
                cell.setCellValue("0");
                cellNum = cellNum + 1;
            }
        } else {
            // 加班一階 二階 三階
            for (int i : resultHours) {
                cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, yellowBack);
                cell.setCellValue(Integer.toString(i));
                cellNum = cellNum + 1;
            }
        }

        // 加班補修
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, yellowBack);
        String restOrMoney = data.getRestOrMoney() == null ? "" : data.getRestOrMoney();
        if (restOrMoney.equals("補休")) {
            cell.setCellValue(String.valueOf(Arrays.stream(resultHours).sum()));
        } else cell.setCellValue("");
        cellNum = cellNum + 1;


        // 空白
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue("");
        cellNum = cellNum + 1;

        // 承認天數(創建時,先設定為空白,等全部資料好了之後 同個人的會合併計算
        cell = CusCell.createCellWithAlignment(workbook, row, cellNum + offset, false);
        cell.setCellValue("");
    }

    /**
     * 如果承認工時為6,7,8 就算一天八小時. 其他時數為多少就是多少
     *
     * @param
     * @return
     */
    private static int[] returnTotalTime(int admitHr) {

        int[] resultHours = {0, 0, 0};
        int oneDayHour = 8;

        // 如果不在6~8 就把onedayHour 設定為admitHr
        if (admitHr == 3 || admitHr == 4) {
            oneDayHour = 4;
        } else if (admitHr < 6 || admitHr > 8) {
            oneDayHour = admitHr;
        }

        if (oneDayHour >= 2) {
            oneDayHour = oneDayHour - 2;
            resultHours[0] = 2;
        } else {
            resultHours[0] = oneDayHour;
            return resultHours;
        }

        if (oneDayHour >= 6) {
            oneDayHour = oneDayHour - 6;
            resultHours[1] = 6;
        } else {
            resultHours[1] = oneDayHour;
            return resultHours;
        }

        resultHours[2] = oneDayHour;

        return resultHours;
    }

    private static int getAdmitHours(Row row) {

        Cell cell1 = row.getCell(9);
        Cell cell2 = row.getCell(10);
        Cell cell3 = row.getCell(11);
        Cell cell4 = row.getCell(12);

        int baseForRestHour = cell4 == null ? 0 : !cell4.getStringCellValue().equals("") ? Integer.parseInt(cell4.getStringCellValue()) : 0;

        if (cell1 != null && cell2 != null && cell3 != null) {
            return baseForRestHour + Integer.parseInt(cell1.getStringCellValue()) + Integer.parseInt(cell2.getStringCellValue())
                    + Integer.parseInt(cell3.getStringCellValue());
        }

        return 0;
    }

    /**
     * 合併編號部分( 依照名字 )
     *
     * @param sheet
     * @param startOfTable2BodyRowNum
     * @param lastRowNumofTable2
     */
    private static void handleTable2Idn(Sheet sheet, int startOfTable2BodyRowNum, int lastRowNumofTable2) {
        int nameColumnIndex = 2;                        // 名稱欄位 在第2行

        // basic needs
        int idn = 1;
        boolean getNameTime = true;
        Cell baseCell;
        Cell baseIdnCell = null;
        String baseName = "";
        int baseIndex = 0;
        Cell nextCell;
        String nextName;

        try {
            for (int i = startOfTable2BodyRowNum; i < lastRowNumofTable2; i++) {
                if (getNameTime) { // 取得名稱
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        baseCell = row.getCell(nameColumnIndex);        // 新的第一個不重複名稱的cell
                        if (baseCell != null) {
                            baseIdnCell = row.getCell(0);            // 該名稱之編號
                            baseName = baseCell.getStringCellValue(); // 該名稱
                            baseIndex = i;                            // 紀錄該rowIndex（合併時會用到)
                            getNameTime = false;                    // 關閉找名字開關( 會在找到下一個不重複名字的時候開啟)
                        }
                    }
                }

                if (!getNameTime) {                                    // 找到不同名稱之前,不會將開關重新開啟
                    if ((i + 1) < lastRowNumofTable2) {                    // 如果下一個row還未超出最大行數
                        Row row = sheet.getRow(i + 1);
                        if (row != null) {
                            nextCell = row.getCell(nameColumnIndex);  // baseCell 的下一列的名稱cell
                            if (nextCell != null) {
                                nextName =
                                        nextCell.getStringCellValue();// 取得下一行名字,看名字是否相同

                                if (!nextName.equals(baseName)) {    // 名字不相同
                                    // 代表這個名稱只有一筆資料
                                    if (i == baseIndex) // 不用合併 因為只有一格
                                        baseIdnCell.setCellValue(idn++); // 設定baseIdnCell編號號碼即可
                                    else { // 不只有一格,要做合併
                                        CellRangeAddress cra = new CellRangeAddress(baseIndex, i, 0, 0);
                                        sheet.addMergedRegion(cra);       // 合併（從baseIndex到目前i的index）
                                        baseIdnCell.setCellValue(idn++); // 設定編號號碼, 都設定到baseIndex那個cell去
                                    }

                                    getNameTime = true;                //把找名稱開關重新開啟
                                }

                            }
                        }
                    } else { // 如果已經是最後一行了
                        if (baseIndex == i) baseIdnCell.setCellValue(idn); // 只有一格,不需合併
                        else {                                             // 不只一格,需要合併
                            try {
                                CellRangeAddress cra = new CellRangeAddress(baseIndex, i, 0, 0);
                                sheet.addMergedRegion(cra);                // 合併 baseIndex 到 i 的cell
                            } catch (Exception e) {
                                System.out.println("在合併儲存格,範圍沒有包含2個以上的cells導致,當加班筆數少於一定數量會有這個例外,如果結果正確可以忽略");
                            }

                            baseIdnCell.setCellValue(idn);
                        }
                    }
                }

            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }


    /**
     * 使用方式以及承認天數核定
     *
     * @param sheet
     * @param startOfTable2BodyRowNum
     * @param lastRowNumofTable2
     */
    private static void handleTable2UsewaysNAdmitDays(Sheet sheet, int startOfTable2BodyRowNum, int lastRowNumofTable2) {

        int nameColumnIndex = 2;
        int restOrMoneyColumnIndex = 8;
        int admitHoursColumnIndex = 14;

        // basic needs
        boolean getNameTime = true;
        Cell baseCell;
        Cell baseAdmitDaysCell = null;
        Cell baseRestOrMoneyCell = null;
        String baseRestOrMoney = "";
        String baseName = "";
        int baseIndex = 0;
        Cell nextCell;
        String nextName;
        String nextRestOrMoney;
        int admitHours = 0;
        try {
            for (int i = startOfTable2BodyRowNum; i < lastRowNumofTable2; i++) {
                if (getNameTime) {                                            // 取得名稱啟動
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        baseCell = row.getCell(nameColumnIndex);                // 取名稱那個欄位的cell
                        if (baseCell != null) {
                            baseAdmitDaysCell = row.getCell(admitHoursColumnIndex); // 這個base的承認天數cell
                            baseRestOrMoneyCell = row.getCell(restOrMoneyColumnIndex); // Base的使用方式cell
                            baseRestOrMoney = baseRestOrMoneyCell.getStringCellValue();// 這個名稱這個行數的 使用方式
                            baseName = baseCell.getStringCellValue();        // 取得名稱
                            baseIndex = i;                                    // 紀錄該rowIndex
                            admitHours = getAdmitHours(row);                // 取得該承認小時

                            getNameTime = false;                            // 關閉找名字開關
                        }
                    }
                }

                if (!getNameTime) {                                            // 找到不同名稱之前,不會將開關重新開啟
                    if ((i + 1) < lastRowNumofTable2) {                            // 如果下一個row還未超出最大行數
                        Row row = sheet.getRow(i + 1);                        // 取得下一行
                        if (row != null) {
                            nextCell = row.getCell(nameColumnIndex);            // 取得下一行的名稱cell
                            if (nextCell != null) {
                                nextName = nextCell.getStringCellValue();    // 取得下一行的名稱
                                nextRestOrMoney = row.getCell(restOrMoneyColumnIndex).getStringCellValue();
                                if (nextName.equals(baseName) && nextRestOrMoney.equals(baseRestOrMoney)) {
                                    // 名字相同 && 使用方式也相同
                                    admitHours = admitHours + getAdmitHours(row); // 把承認天數加起來
                                } else { // 名字不相同
                                    mergeModule(i, baseIndex, sheet, baseRestOrMoneyCell, baseAdmitDaysCell, baseRestOrMoney, admitHours);
                                    //把getNameTime重新開啟
                                    getNameTime = true;
                                }
                            }
                        }
                    } // end if()
                    // 如果已經是最後一行了
                    else {
                        try {
                            mergeModule(i, baseIndex, sheet, baseRestOrMoneyCell, baseAdmitDaysCell, baseRestOrMoney, admitHours);
                        } catch (Exception e) {
                            System.out.println(e.getMessage());
                            System.out.println("在合併儲存格,範圍沒有包含2個以上的cells導致,當加班筆數少於一定數量會有這個例外,如果結果正確可以忽略");
                        }
                    }
                }

            }
        } catch (Exception e) {
            // TODO: handle exception
        }

    }

    private static void mergeModule(int i, int baseIndex, Sheet sheet, Cell baseRestOrMoneyCell, Cell baseAdmitDaysCell, String baseRestOrMoney, int admitHours) {
        if (i != baseIndex) { // 這個名稱只有一筆資料
            // 代表 要做合併了 從baseIndex到nextIndex
            CellRangeAddress cra = new CellRangeAddress(baseIndex, i, 8, 8);
            sheet.addMergedRegion(cra);

            // 代表 要做合併了 從baseIndex到nextIndex
            cra = new CellRangeAddress(baseIndex, i, 14, 14);
            sheet.addMergedRegion(cra);
        }

        baseRestOrMoneyCell.setCellValue(baseRestOrMoney);
        baseAdmitDaysCell.setCellValue(returnAdmitDays(admitHours));
    }


    private static String returnAdmitDays(int inhours) {

        int days = inhours / 8;
        int hours = inhours % 8;

        return days + "天" + hours + "時";
    }

    // Sheet2 Table1

    /**
     * 生成sheet表2，並寫入第一行數據（列頭）and 相關說明 更新日期等
     *
     * @param workbook 工作簿對象
     * @return 已經寫入列頭的Sheet
     */
    private static Sheet buildDataSheet2(Workbook workbook, List<PinkPojo> pinkPojos) {

        Sheet sheet = workbook.createSheet("震旦雲打卡明細"); // createSheet裡面放參數 可指定工作表名稱
        // 設置 Column 寬度
        for (int i = 0; i < PINK_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 5000);
        }
        // 設置默認行高
        sheet.setDefaultRowHeight((short) 400);

        convertPinkPojosToRow(pinkPojos, workbook, sheet);

        return sheet;
    }

    /**
     * 將有在excel上並且顯示粉紅色的（假日加班）, 但沒有word加班單的顯示在我們的sheet2
     *
     * @param pinkPojos
     * @param workbook
     * @param sheet
     */
    private static void convertPinkPojosToRow(List<PinkPojo> pinkPojos, Workbook workbook, Sheet sheet) {
        int startRow = 0;
        CellStyle cellStyle = CusCellStyle.buildHeadCellStyle(sheet.getWorkbook());
        // 創建第二個table的表頭
        Row head = sheet.createRow(startRow++);
        for (int i = 0; i < PINK_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(PINK_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }

        for (PinkPojo pinkPojo : pinkPojos) {
            int cellNum = 0;
            Cell cell;
            Row row = sheet.createRow(startRow++);

            // 編號
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
            cell.setCellValue(row.getRowNum());

            // 日期
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
            cell.setCellValue(pinkPojo.getDate() == null ? "" : pinkPojo.getDate());

            // 員工
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
            cell.setCellValue(pinkPojo.getEmployee() == null ? "" : pinkPojo.getEmployee());

            // 上班
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
            cell.setCellValue(pinkPojo.getOnTime() == null ? "" : pinkPojo.getOnTime());

            // 下班
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum++, false);
            cell.setCellValue(pinkPojo.getOffTime() == null ? "" : pinkPojo.getOffTime());

            // 缺卡內容
            cell = CusCell.createCellWithAlignment(workbook, row, cellNum, false);
            cell.setCellValue(pinkPojo.getMissContent() == null ? "" : pinkPojo.getMissContent());
        }
    }

    /**
     * @param admitTime
     * @return
     */
    private static Boolean redRule(String admitTime) {
        int admitHour = Integer.parseInt(admitTime);
        if (admitHour > 9)
            return true;

        switch (Integer.parseInt(admitTime)) {
            case 1:
            case 2:
            case 5:
            case 9:
                return true;
            default:
                return false;
        }
    }

}