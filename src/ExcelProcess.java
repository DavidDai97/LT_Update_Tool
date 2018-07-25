import jxl.*;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import javax.swing.*;
import java.io.*;
import java.lang.Number;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.concurrent.locks.LockSupport;

public class ExcelProcess {
    private static SimpleDateFormat myFormat = new SimpleDateFormat("yyyyMMdd");
    private static WritableCellFormat titleFormat;
    private static WritableCellFormat normalFormat;
    private static WritableCellFormat expiredFormat;
    private static WritableCellFormat noneFormat;
    public static int currentProcess = 0;
    public static int totalRowNums = 0;

    public static void initializeFormat(){
        try{
            WritableFont titleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,false);
            titleFormat = new WritableCellFormat(titleFont);
            titleFormat.setAlignment(Alignment.CENTRE);
            titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            WritableFont myFont = new WritableFont(WritableFont.ARIAL,10, WritableFont.NO_BOLD, false);
            expiredFormat = new WritableCellFormat(myFont);
            noneFormat = new WritableCellFormat(myFont);
            normalFormat = new WritableCellFormat(myFont);
            normalFormat.setAlignment(Alignment.CENTRE);
            normalFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            expiredFormat.setBackground(Colour.RED);
            expiredFormat.setAlignment(Alignment.CENTRE);
            expiredFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            noneFormat.setBackground(Colour.YELLOW);
            noneFormat.setAlignment(Alignment.CENTRE);
            noneFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        }
        catch(WriteException e){
            System.out.println("Err: 5, Initialize Error.");
        }
    }
    private static Queue<Sheet> getSheetNum(Workbook wb, String text){
        int sheet_size = wb.getNumberOfSheets();
        Queue<Sheet> results = new LinkedList<>();
        for (int index = 0; index < sheet_size; index++){
            Sheet dataSheet = wb.getSheet(index);
            if(dataSheet.getName().contains(text)){
                results.add(dataSheet);
                if(text.equals("Sheet 1")) {
                    totalRowNums += dataSheet.getRows();
                }
            }
        }
        totalRowNums = totalRowNums;
        return results;
    }
    private static void setProgressBar(){
        currentProcess++;
        MainGUI.processBar.setValue((int)(((double)currentProcess/(double)totalRowNums)*100));
    }
    private static void processDataHelper(File poFile, File masterFile) throws Exception{
        InputStream poFileIs = new FileInputStream(poFile.getAbsolutePath());
        InputStream masterFileIs = new FileInputStream(masterFile.getAbsolutePath());
        Workbook poWb = Workbook.getWorkbook(poFileIs);
        Workbook masterWb = Workbook.getWorkbook(masterFileIs);
        Queue<Sheet> poDataSheets = getSheetNum(poWb, "Sheet 1");
        Queue<Sheet> masterDataSheets = getSheetNum(masterWb, "Item_Master");
        Sheet[] masterSheet = masterDataSheets.toArray(new Sheet[0]);
        int orderDateColIdx, deliveryDateColIdx, supplierColIdx, orderTypeColIdx, buyerColIdx, itemDescriptionColIdx, priceColIdx, currencyColIdx, itemNumIdx, processLeadTimeIdx;
        int quantityColIdx, moqColIdx;
        while(!poDataSheets.isEmpty()) {
            Sheet currDataSheet = poDataSheets.poll();
            orderDateColIdx = currDataSheet.findCell("Po Line Creation Date").getColumn();
            deliveryDateColIdx = currDataSheet.findCell("Transaction Date").getColumn();
            supplierColIdx = currDataSheet.findCell("Supplier Name").getColumn();
            orderTypeColIdx = currDataSheet.findCell("Line Type").getColumn();
            buyerColIdx = currDataSheet.findCell("Buyer Name").getColumn();
            itemDescriptionColIdx = currDataSheet.findCell("Po Item Description").getColumn();
            priceColIdx = currDataSheet.findCell("Po Unit Price").getColumn();
            currencyColIdx = currDataSheet.findCell("Po Currency").getColumn();
            itemNumIdx = currDataSheet.findCell("Item Num").getColumn();
            processLeadTimeIdx = masterDataSheets.peek().findCell("Processing Lead Time").getColumn();
            quantityColIdx = currDataSheet.findCell("Quantity Received").getColumn();
            moqColIdx = masterDataSheets.peek().findCell("Minimum Order Qty").getColumn();
            int rowNum = currDataSheet.getRows();
            for(int i = 1; i < rowNum; i++){
                System.out.println("Row: " + i + ", " + totalRowNums + ", " + (int)(((double)currentProcess/(double)totalRowNums)*100));
                setProgressBar();
                if (!currDataSheet.getCell(orderTypeColIdx, i).getContents().equals("Goods")) {
                    continue;
                }
                String currItemNum = currDataSheet.getCell(itemNumIdx, i).getContents();
                DateCell currOrderDateCell = (DateCell) currDataSheet.getCell(orderDateColIdx, i);
                Date currOrderDate_temp = currOrderDateCell.getDate();
                String currOrderDate = myFormat.format(currOrderDate_temp);
                DateCell currDeliveryDateCell = (DateCell) currDataSheet.getCell(deliveryDateColIdx, i);
                Date currDeliveryDate_temp = currDeliveryDateCell.getDate();
                String currDeliveryDate = myFormat.format(currDeliveryDate_temp);
                double currPrice = ((NumberCell) currDataSheet.getCell(priceColIdx, i)).getValue();
                int quantityReceived = (int)(((NumberCell) currDataSheet.getCell(quantityColIdx, i)).getValue());
                if(MainGUI.itemNumList.contains(currItemNum)){
                    Item currItem = MainGUI.itemMap.get(currItemNum);
                    currItem.addPurchaseTime();
                    currItem.setAverageLeadTime(getDutyDays(currOrderDate, currDeliveryDate));
                    currItem.setPrice(currPrice);
                    currItem.addQuantity(quantityReceived);
                    continue;
                }
                String currSupplier = currDataSheet.getCell(supplierColIdx, i).getContents();
                String currDescription = currDataSheet.getCell(itemDescriptionColIdx, i).getContents();
                String currBuyer = currDataSheet.getCell(buyerColIdx, i).getContents();
                String currCurrency = currDataSheet.getCell(currencyColIdx, i).getContents();
                int processLeadtime = 0;
                int currMoq = 0;
                for(int j = 0; j < masterSheet.length; j++){
                    Sheet currMasterSheet = masterSheet[j];
                    if(currMasterSheet.findCell(currItemNum) != null){
                        int currItemRow = currMasterSheet.findCell(currItemNum).getRow();
                        if(currMasterSheet.getCell(moqColIdx, currItemRow).getContents() == ""){
                            currMoq = 0;
                        }
                        else{
                            currMoq = (int)((NumberCell)currMasterSheet.getCell(moqColIdx, currItemRow)).getValue();
                        }
                        if(currMasterSheet.getCell(processLeadTimeIdx, currItemRow).getContents() == ""){
                            processLeadtime = 0;
                        }
                        else {
                            processLeadtime = (int) ((NumberCell) currMasterSheet.getCell(processLeadTimeIdx, currItemRow)).getValue();
                        }
                        break;
                    }
                }
                Item currItem = new Item(currBuyer, getDutyDays(currOrderDate, currDeliveryDate), currCurrency, currPrice, currSupplier, processLeadtime, currDescription, currMoq, quantityReceived);
                MainGUI.itemNumList.add(currItemNum);
                MainGUI.itemMap.put(currItemNum, currItem);
            }
        }
    }
    private static int writeCnt = 0;
    private static void outputFile(){
        int itemNumCol = 0;
        int itemDescriptionCol = 1;
        int itemSupplierCol = 2;
        int itemLeadTimeCol = 9;
        int currencyCol = 4;
        int itemPriceCol = 3;
        int itemProcessLTCol = 10;
        int itemBuyerCol = 5;
        int varianceCol = 11;
        int totalQuantityCol = 6;
        int purchaseTimesCol = 7;
        int itemMOQCol = 8;
        try {
            String outputFilePath = "../Output/LT Comparison_" + getTodayDate() + ".xls";
            WritableWorkbook outputFile = Workbook.createWorkbook(new File(outputFilePath));
            WritableSheet catlogSheet = outputFile.createSheet("Comparison", 0);
            jxl.write.Label itemNameTitle = new jxl.write.Label(itemNumCol, 0, "Item Num", titleFormat);
            catlogSheet.addCell(itemNameTitle);
            catlogSheet.setColumnView(itemNumCol, 15);
            jxl.write.Label partNumTitle = new jxl.write.Label(itemDescriptionCol, 0, "Item Description", titleFormat);
            catlogSheet.addCell(partNumTitle);
            catlogSheet.setColumnView(itemDescriptionCol, 50);
            jxl.write.Label brandTitle = new jxl.write.Label(itemProcessLTCol, 0, "Processing Lead Time", titleFormat);
            catlogSheet.addCell(brandTitle);
            catlogSheet.setColumnView(itemProcessLTCol, 20);
            jxl.write.Label supplierTitle = new jxl.write.Label(itemSupplierCol, 0, "Supplier", titleFormat);
            catlogSheet.addCell(supplierTitle);
            catlogSheet.setColumnView(itemSupplierCol, 30);
            jxl.write.Label leadTimeTitle = new jxl.write.Label(itemLeadTimeCol, 0, "Lead Time", titleFormat);
            catlogSheet.addCell(leadTimeTitle);
            catlogSheet.setColumnView(itemLeadTimeCol, 10);
            jxl.write.Label currencyTitle = new jxl.write.Label(currencyCol, 0, "Currency", titleFormat);
            catlogSheet.addCell(currencyTitle);
            catlogSheet.setColumnView(currencyCol, 9);
            jxl.write.Label priceTitle = new jxl.write.Label(itemPriceCol, 0, "Item Price", titleFormat);
            catlogSheet.addCell(priceTitle);
            catlogSheet.setColumnView(itemPriceCol, 12);
            jxl.write.Label buyerTitle = new jxl.write.Label(itemBuyerCol, 0, "Buyer", titleFormat);
            catlogSheet.addCell(buyerTitle);
            catlogSheet.setColumnView(itemBuyerCol, 15);
            jxl.write.Label varianceTitle = new jxl.write.Label(varianceCol, 0, "Lead Time Variance", titleFormat);
            catlogSheet.addCell(varianceTitle);
            catlogSheet.setColumnView(varianceCol, 9);
            jxl.write.Label totalQuantityTitle = new jxl.write.Label(totalQuantityCol, 0, "Received Quantity", titleFormat);
            catlogSheet.addCell(totalQuantityTitle);
            catlogSheet.setColumnView(totalQuantityCol, 13);
            jxl.write.Label purchaseTimesTitle = new jxl.write.Label(purchaseTimesCol, 0, "Purchase Frequency", titleFormat);
            catlogSheet.addCell(purchaseTimesTitle);
            catlogSheet.setColumnView(purchaseTimesCol, 14);
            jxl.write.Label moqTitle = new jxl.write.Label(itemMOQCol, 0, "MOQ", titleFormat);
            catlogSheet.addCell(moqTitle);
            catlogSheet.setColumnView(itemMOQCol, 5);
            for(int i = 0; i < MainGUI.itemNumList.size(); i++){
                Item currItem = MainGUI.itemMap.get(MainGUI.itemNumList.get(i));
                WritableCellFormat currFormat = normalFormat;
                if(currItem.getColor() == 0){
                    currFormat = normalFormat;
                }
                else if(currItem.getColor() == 1){
                    currFormat = expiredFormat;
                }
                else if(currItem.getColor() == 2){
                    currFormat = noneFormat;
                }
                // itemNum, itemDescription, itemLeadTime, itemProcessLT, itemSupplier, itemPrice, itemCurrency, itemBuyer
                Label itemNumLabel = new Label(itemNumCol, i+1, MainGUI.itemNumList.get(i), normalFormat);
                catlogSheet.addCell(itemNumLabel);
                Label itemDescriptionLabel = new Label(itemDescriptionCol, i+1, currItem.getDescription(), normalFormat);
                catlogSheet.addCell(itemDescriptionLabel);
                jxl.write.Number currLeadTimeLabel = new jxl.write.Number(itemLeadTimeCol, i+1, currItem.getAverageLeadTime(), normalFormat);
                catlogSheet.addCell(currLeadTimeLabel);
                jxl.write.Number currProcessLTLabel = new jxl.write.Number(itemProcessLTCol, i+1, currItem.getProcessLeadTime(), normalFormat);
                catlogSheet.addCell(currProcessLTLabel);
                Label itemSupplierLabel = new Label(itemSupplierCol, i+1, currItem.getSupplier(), normalFormat);
                catlogSheet.addCell(itemSupplierLabel);
                jxl.write.Number currPriceLabel = new jxl.write.Number(itemPriceCol, i+1, currItem.getPrice(), normalFormat);
                catlogSheet.addCell(currPriceLabel);
                Label itemCurrencyLabel = new Label(currencyCol, i+1, currItem.getCurrency(), normalFormat);
                catlogSheet.addCell(itemCurrencyLabel);
                Label itemBuyerLabel = new Label(itemBuyerCol, i+1, currItem.getBuyer(), normalFormat);
                catlogSheet.addCell(itemBuyerLabel);
                jxl.write.Number currVarianceLabel = new jxl.write.Number(varianceCol, i+1, currItem.getDifference(), currFormat);
                catlogSheet.addCell(currVarianceLabel);
                // Total Quantity, Purchase Times, MOQ
                jxl.write.Number currTotalQuantityLabel = new jxl.write.Number(totalQuantityCol, i+1, currItem.getTotalQuantity(), normalFormat);
                catlogSheet.addCell(currTotalQuantityLabel);
                jxl.write.Number currPurchaseTimesLabel = new jxl.write.Number(purchaseTimesCol, i+1, currItem.getPurchaseTime(), normalFormat);
                catlogSheet.addCell(currPurchaseTimesLabel);
                if(currItem.getMOQ() != 0){
                    jxl.write.Number currMOQLabel = new jxl.write.Number(itemMOQCol, i+1, currItem.getMOQ(), normalFormat);
                    catlogSheet.addCell(currMOQLabel);
                }
            }
            outputFile.write();
            outputFile.close();
            writeCnt = 0;
        }
        catch (Exception e){
            System.out.println(e.toString());
            writeCnt++;
            if(writeCnt < 5){
                initializeFormat();
                outputFile();
            }
        }
    }
    private static String getTodayDate(){
        Calendar myCalendar = Calendar.getInstance();
        return myFormat.format(myCalendar.getTime());
    }
    private static int getDutyDays(String strStartDate,String strEndDate) {
        Date startDate;
        Date endDate;
        try {
            startDate=myFormat.parse(strStartDate);
            endDate = myFormat.parse(strEndDate);
        }catch(ParseException e) {
            System.out.println("Format error");
            e.printStackTrace();
            return 0;
        }
        int result = 0;
        int days = (int) ((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000)) + 1;
        int weeks = days / 7;
        if (days % 7 == 0) {
            result = days - 2 * weeks;
        }
        else{
            Calendar begCalendar = Calendar.getInstance();
            Calendar endCalendar = Calendar.getInstance();
            begCalendar.setTime(startDate);
            endCalendar.setTime(endDate);
            int beg = begCalendar.get(Calendar.DAY_OF_WEEK);
            int end = endCalendar.get(Calendar.DAY_OF_WEEK);
            if(beg > end){
                result = days - 2 * (weeks + 1);
            }
            else if(beg < end){
                if(end == 7){
                    result = days - 2 * weeks - 1;
                }
                else{
                    result = days - 2 * weeks;
                }
            }
            else{
                if(beg == 1 || beg == 7){
                    result = days - 2 * weeks - 1;
                }
                else{
                    result = days - 2 * weeks;
                }
            }
        }
        return result;
    }


    public static void processData(String poFileName, String masterFileName) throws Exception{
        String poFilePath = "../Source/" + poFileName;
        String masterFilePath = "../Source/" + masterFileName;
        File poFile = new File(poFilePath);
        File masterFile = new File(masterFilePath);
        processDataHelper(poFile, masterFile);
        outputFile();
    }
}
