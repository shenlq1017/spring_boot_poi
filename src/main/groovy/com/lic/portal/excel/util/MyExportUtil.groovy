package com.lic.portal.excel.util

import cn.afterturn.easypoi.excel.ExcelExportUtil
import cn.afterturn.easypoi.excel.entity.ExportParams
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity
import cn.afterturn.easypoi.excel.entity.vo.BaseEntityTypeConstants
import cn.afterturn.easypoi.excel.entity.vo.PoiBaseConstants
import cn.afterturn.easypoi.excel.export.ExcelExportService
import cn.afterturn.easypoi.excel.export.base.BaseExportService
import cn.afterturn.easypoi.excel.export.base.ExportCommonService
import cn.afterturn.easypoi.excel.export.styler.IExcelExportStyler
import cn.afterturn.easypoi.exception.excel.ExcelExportException
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler
import cn.afterturn.easypoi.handler.inter.IExcelDictHandler
import cn.afterturn.easypoi.util.PoiExcelGraphDataUtil
import com.lic.portal.excel.model.CellMapping
import com.lic.portal.excel.model.CellMerge
import com.lic.portal.excel.model.ExportEntity
import com.lic.portal.excel.model.ExportParamsFoot
import com.lic.portal.excel.model.ExportParamsHeader
import org.apache.commons.lang3.StringUtils
import org.apache.commons.lang3.builder.ReflectionToStringBuilder
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Drawing
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference

class MyExportUtil extends BaseExportService {

    protected IExcelDataHandler dataHandler;
    protected IExcelDictHandler dictHandler;
    List<String> needHandlerList;

    // 最大行数,超过自动多Sheet
    private static int MAX_NUM = 60000;

    /**
     * 导出
     * @param exportEntities
     * @param excelType
     * @param dataSize
     * @return
     */
    static Workbook exportExcel(List<ExportEntity> exportEntities, ExcelType excelType,int dataSize) {
        Workbook workbook = ExcelExportUtil.getWorkbook(excelType,dataSize);;
        for(ExportEntity exportEntity : exportEntities) {
            new MyExportUtil().createSheetForMap(workbook, exportEntity.getEntity(),exportEntity.getExportParamsFoot(), (List<ExcelExportEntity>) exportEntity.getEntityList(), (Collection) exportEntity.getDataSet());
        }
        return workbook;
    }

    def createSheetForMap(Workbook workbook, ExportParamsHeader entity,ExportParamsFoot exportParamsFoot,List<ExcelExportEntity> entityList, Collection<?> dataSet) {
        if (workbook == null || entity == null || entityList == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        if (ExcelType.XSSF.equals(entity.getType())) {
            MAX_NUM = 1000000;
        }
        if (entity.getMaxNum() > 0) {
            MAX_NUM = entity.getMaxNum();
        }
        Sheet sheet = null;
        try {
            sheet = workbook.createSheet(entity.getSheetName());
        } catch (Exception e) {
            // 重复遍历,出现了重名现象,创建非指定的名称Sheet
            sheet = workbook.createSheet();
        }
        insertDataToSheet(workbook, entity,exportParamsFoot, entityList, dataSet, sheet);
    }

    protected int createHeaderAndTitle(ExportParamsHeader entity, Sheet sheet, Workbook workbook,
                                       List<ExcelExportEntity> excelParams, Map<String,CellMapping> cellMappingMap,int fieldLength) {
        int rows = 0;
        if (entity.getTitle() != null) {
            rows += createTitle2Row(entity, sheet, workbook, fieldLength-1);
        }
        rows += createHeaderRowN(entity, sheet, workbook, rows, excelParams,cellMappingMap);
        return rows;
    }

    /**
     * 拿到合并的总表格长度
     *
     */
    int titleWidth(List<ExcelExportEntity> excelParams) {
        int num = 0;
        for(ExcelExportEntity entity : excelParams) {
            if(entity.getList()==null || entity.getList().isEmpty()) {
                if(entity.needMerge) {
                    num +=entity.getMergeRely().length>0?entity.getMergeRely()[0]==0?1:entity.getMergeRely()[0]:1;
                }else{
                    num++;
                }
            }else {

                int childNum = 1 ;
                ExcelExportEntity entityChild = entity.getList().get(0);
                if(entityChild.getName()==null) {
                    childNum = 0;
                }else {
                    int childTempNum = entityChild.getMergeRely().length > 0 ? entityChild.getMergeRely()[0] == 0 ? 1 : entityChild.getMergeRely()[0] + 1: 1;
                    if (childTempNum > childNum) {
                        childNum = childTempNum;
                    }
                }
                num+=childNum;
            }
        }
        return num;
    }

    /**
     * 创建 表头改变
     */
    int createTitle2Row(ExportParamsHeader entity, Sheet sheet, Workbook workbook,
                               int fieldWidth) {

        Row row = sheet.createRow(0);
        row.setHeight((short) (50*20));
        CellStyle titleStyle = getExcelExportStyler().getHeaderStyle(entity.getHeaderColor());
        titleStyle = MyExcelHeaderUtil.setTitleStyle(sheet,workbook,entity,titleStyle)
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER)
        createStringCell(row, 0, entity.getTitle(),titleStyle, null);
//        for (int i = 1; i <= fieldWidth; i++) {
//            createStringCell(row, i, "",
//                    getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
//        }
        CellRangeAddress cellRangeAddressTitle = new CellRangeAddress(0, 0, 0, fieldWidth)
        sheet.addMergedRegion(cellRangeAddressTitle);
//        MyCellStyleUtil.setBorderForMergeCell(1,cellRangeAddressTitle,sheet,workbook);
        if (entity.getSecondTitle() != null) {
            row = sheet.createRow(1);
            row.setHeight(entity.getSecondTitleHeight());

//            for (int i = 1; i <= fieldWidth; i++) {
//                createStringCell(row, i, "",
//                        getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
//            }
            String[] secondTitleStrArr = entity.getSecondTitle().split("&&");
            int secondTitleStrArrLen = secondTitleStrArr.length;
            for(int i =0;i<secondTitleStrArrLen;i++) {
                CellStyle style = workbook.createCellStyle();

                int cellxS=i==0?0:(int) (i * (fieldWidth/2)) +1;
                int cellxE = i==secondTitleStrArrLen-1?fieldWidth:(int) ((i+1) * (fieldWidth/2)-1)+1;

                style.setAlignment(i == 0?CellStyle.ALIGN_LEFT:CellStyle.ALIGN_RIGHT);
                createStringCell(row, cellxS,secondTitleStrArr[i], style, null);

                CellRangeAddress cellRangeAddressSecondTitle = new CellRangeAddress(1, 1, cellxS, cellxE)
                sheet.addMergedRegion(cellRangeAddressSecondTitle);
                MyCellStyleUtil.onBorderForMergeCell((int) CellStyle.BORDER_NONE, cellRangeAddressSecondTitle, sheet, workbook);
            }
            return 2;
        }
        return 1;
    }

    /**
     * 创建表头
     */
    private int createHeaderRow(ExportParams title, Sheet sheet, Workbook workbook, int index,
                                List<ExcelExportEntity> excelParams) {
        Row row = sheet.createRow(index);
        int rows = getRowNums(excelParams);
        row.setHeight(title.getHeaderHeight());
        Row listRow = null;
        if (rows >= 2) {
            listRow = sheet.createRow(index + 1);
            listRow.setHeight(title.getHeaderHeight());
        }
        int cellIndex = 0;
        int groupCellLength = 0;
        CellStyle titleStyle = getExcelExportStyler().getTitleStyle(title.getColor());
        int exportFieldTitleSize = excelParams.size();
        for (int i = 0; i < exportFieldTitleSize; i++) {
            ExcelExportEntity entity = excelParams.get(i);
            // 加入换了groupName或者结束就，就把之前的那个换行
            if (StringUtils.isBlank(entity.getGroupName()) || !entity.getGroupName().equals(excelParams.get(i - 1).getGroupName())) {
                if (groupCellLength > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex - groupCellLength, cellIndex - 1));
                }
                groupCellLength = 0;
            }
            if (StringUtils.isNotBlank(entity.getGroupName())) {
                createStringCell(row, cellIndex, entity.getGroupName(), titleStyle, entity);
                createStringCell(listRow, cellIndex, entity.getName(), titleStyle, entity);
                groupCellLength++;
            } else if (StringUtils.isNotBlank(entity.getName())) {
                createStringCell(row, cellIndex, entity.getName(), titleStyle, entity);
            }
            if (entity.getList() != null) {
                List<ExcelExportEntity> sTitel = entity.getList();
                if (StringUtils.isNotBlank(entity.getName()) && sTitel.size() > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex, cellIndex + sTitel.size() - 1));
                }

                int sTitelSize = sTitel.size();
                for (int j = 0; j < sTitelSize; j++) {
                    createStringCell(rows >= 2 ? listRow : row, cellIndex, sTitel.get(j).getName(),
                            titleStyle, entity);
                    cellIndex++;
                }
                cellIndex--;
            } else if (rows >= 2 && StringUtils.isBlank(entity.getGroupName())) {
                createStringCell(listRow, cellIndex, "", titleStyle, entity);
                sheet.addMergedRegion(new CellRangeAddress(index, index + 1, cellIndex, cellIndex));
            }
            cellIndex++;
        }
        if (groupCellLength > 1) {
            sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex - groupCellLength, cellIndex - 1));
        }
        return rows;
    }

    /**
     * 标题栏
     * @param title
     * @param sheet
     * @param workbook
     * @param index
     * @param excelParams
     * @return
     */
    private int createHeaderRowN(ExportParams title, Sheet sheet, Workbook workbook, int index,
                                 List<ExcelExportEntity> excelParams, Map<String,CellMapping> cellMappingMap) {
//        index +=1;
        int rows = getRowNums(excelParams);
        List<Row> rowsList = new ArrayList<>();
        //先建好所有行
        for(int i = index;i<rows+index;i++ ){

            Row row = sheet.createRow(i);
            row.setHeight((short) (20 * 20));
            rowsList.add(row);
        }


        int exportFieldTitleSize = excelParams.size();

        //设置初始y值，当横向合并以后,以row为key
        Map<Integer,Integer> sety = new HashMap<>();

        //遍历每个标题实体
        for (int j = 0; j < exportFieldTitleSize; j++) {

            //初始化列号
            int celly = index;

            ExcelExportEntity exportEntityThis = excelParams.get(j);
            ExcelExportEntity exportEntityTemp = exportEntityThis;
            CellStyle titleStyle = getExcelExportStyler().getTitleStyle(title.getColor());
            titleStyle = MyCellStyleUtil.onBorder(titleStyle,(short) 1);
//            titleStyle = MyCellStyleUtil.onFontBold(workbook,titleStyle)
            //拿到标题子项中的下标
            int dataRow = 0;

            //按行填写
            for(int i =index;i<rows+index;i++) {
                //关键...获取当前行号最后记录的列号
                int cellx = sety.get(i) != null ? sety.get(i) : 0;
                //如果是第一行，且没有子项，或者有子项
                if((i == index && exportEntityThis.getList() !=null) || (i>=1 && exportEntityThis.getList() !=null && exportEntityThis.getList().size() > 1)) {
                    if(dataRow>=exportEntityThis.getList().size()){
                        continue;
                    }
                    exportEntityTemp = exportEntityThis.getList().get(dataRow);
                    dataRow++;
                }else if(i>index && exportEntityThis.getList() ==null) {
                    continue;
                }
                cellMappingMap.put(exportEntityTemp.getKey(),new CellMapping(cellx,exportEntityTemp.getType()))
                //如果名称是null则不用创建
                if(exportEntityTemp.getName()==null) {
                    continue;
                }else {
                    //创建表
                    createStringCell(rowsList.get(i - index), cellx, exportEntityTemp.getName(), titleStyle, exportEntityTemp);
                    //需要合并
                    if (exportEntityTemp.needMerge) {
                        int[] cellNums = cellMerge(sheet, exportEntityTemp, cellx, celly,sety,workbook);
                        cellx = cellNums[0];
                        celly = cellNums[1];
                    }else {
                        MyCellUtil.cellWidth(exportEntityTemp,sheet,cellx);
                        sety.put(i, cellx + 1);
                    }
                    celly++;
                }
            }
        }
        return rows;

    }



    int[] cellMerge(Sheet sheet,ExcelExportEntity exportEntityThis,int cellx,int celly,Map<Integer,Integer> sety,Workbook workbook) {
        //拿到横向和纵的数量，0 表示不扩展
        int[] mergerNum = exportEntityThis.mergeRely;
        int rowMerger = 0;
        int columnMerger = 0;
        if(mergerNum.length>0) {
            rowMerger = mergerNum[0];
            if(mergerNum.length>1) {
                columnMerger = mergerNum[1];
            }
        }
        if(rowMerger>0 || columnMerger>0) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(celly,celly+columnMerger,cellx,cellx+rowMerger);
            MyCellStyleUtil.onBorderForMergeCell( 1,cellRangeAddress,sheet,workbook);
            sheet.addMergedRegion(cellRangeAddress);
            for(int i = celly;i<celly+columnMerger+1;i++){
                sety.put(i,cellx+rowMerger+1);
            }
            for (int i = cellx; i < cellx+rowMerger+1; i++) {
                MyCellUtil.cellWidth(exportEntityThis,sheet,i);
            }
            cellx = cellx+rowMerger;
            celly = celly+columnMerger;

        }
        return [cellx,celly];
    }


    /**
     * 拿到有几行表头
     * @param excelParams
     * @return
     */
    int getRowNums(List<ExcelExportEntity> excelParams) {
        int rowNums = 1;
        for (int i = 0; i < excelParams.size(); i++) {
            if(excelParams.get(i).getList() !=null ){
                if(excelParams.get(i).getList().size() > rowNums) {
                    rowNums = excelParams.get(i).getList().size();
                }
            }
        }
        return rowNums;
    }

    def insertDataToSheet(Workbook workbook, ExportParams entity, ExportParamsFoot exportParamsFoot,List<ExcelExportEntity> entityList, Collection<?> dataSet, Sheet sheet) {
        try {
            dataHandler = entity.getDataHandler();
            if (dataHandler != null && dataHandler.getNeedHandlerFields() != null) {
                needHandlerList = Arrays.asList(dataHandler.getNeedHandlerFields());
            }
            dictHandler = entity.getDictHandler();
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) entity.getStyle()
                    .getConstructor(Workbook.class).newInstance(workbook));
            Drawing patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            int fieldLength = titleWidth(excelParams);
            Map<String,CellMapping> cellMappingMap = new HashMap<>()

            int index = entity.isCreateHeadRows()? createHeaderAndTitle(entity, sheet, workbook, excelParams,cellMappingMap,fieldLength) : 0;
            int titleHeight = index;
//            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);
            short rowHeight = entity.getHeight() != 0 ? entity.getHeight() : getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += myCreateCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight,cellMappingMap);
                tempList.add(t);
                if (index >= MAX_NUM) {
                    break;
                }
            }
            if (entity.getFreezeCol() != 0) {
                sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
            }

//            mergeCells(sheet, excelParams, titleHeight);

            its = dataSet.iterator();
            int leSize =  tempList.size()
            for (int i = 0; i < leSize; i++) {
                its.next();
                its.remove();
            }
            // 发现还有剩余list 继续循环创建Sheet
            if (dataSet.size() > 0) {
                createSheetForMap(workbook, entity, entityList, dataSet);
            } else {
                if(exportParamsFoot!=null) {
                    MyExcelFootUtil.creatFoot(index, sheet, workbook, exportParamsFoot, fieldLength);
                }
                // 创建合计信息
                addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
            }

        } catch (Exception e) {
            e.printStackTrace()
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    ExcelExportEntity indexExcelEntity(ExportParams entity) {
        ExcelExportEntity exportEntity = new ExcelExportEntity();
        //保证是第一排
        exportEntity.setOrderNum(Integer.MIN_VALUE);
        exportEntity.setName(entity.getIndexName());
        exportEntity.setWidth(10);
        exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
        return exportEntity;
    }


    /**
     * 复写创建 最主要的 Cells
     */
    public int myCreateCells(Drawing patriarch, int index, Object t,
                           List<ExcelExportEntity> excelParams, Sheet sheet, Workbook workbook,
                           short rowHeight,Map<String,CellMapping> cellMappingMap) {
        try {
            ExcelExportEntity entity;
            Row row = sheet.createRow(index);
            if (rowHeight != -1) {
                row.setHeight(rowHeight);
            }
            int maxHeight = 1, cellNum = 0;
            int indexKey = super.createIndexCell(row, index, excelParams.get(0));
            cellNum += indexKey;
            int paramSize = excelParams.size()

            Map<String,Object> data = (Map<String,Object>) t;
            for(Map.Entry<String,Object> map : data.entrySet()){
                String thisDatakey = map.getKey();
                Object value = null;
                int mergeNum = 0;
                if(CellMerge.MERVE_COLUMN_KEY.equals(thisDatakey)) {
                    List<CellMerge> cellMerges = (List<CellMerge>) map.getValue();
                    for(CellMerge cellMerge : cellMerges ) {
                        thisDatakey = cellMerge.getCellKey()
                        value = cellMerge.getCellVal()
                        int mergeLengthways = cellMerge.getMergeLengthways()
                        int mergeCrosswise = cellMerge.getMergeCrosswise()
                        CellMapping cellMapping = cellMappingMap.get(thisDatakey);
                        cellValSet(cellMapping,cellMappingMap,value,mergeLengthways,mergeCrosswise,row,sheet,workbook,index);
                    }
                }else {
                    CellMapping cellMapping = cellMappingMap.get(thisDatakey);
                    cellValSet(cellMapping,cellMappingMap,map.getValue(),0,0,row,sheet,workbook,index);
                }

            }

        } catch (Exception e) {
            e.printStackTrace()
            LOGGER.error("excel cell export error ,data is :{}", ReflectionToStringBuilder.toString(t));
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
        return 1;

    }

    def cellValSet(CellMapping cellMapping,Map<String,CellMapping> cellMappingMap,
                   Object value,int mergeLengthways,int mergeCrosswise,Row row,Sheet sheet,Workbook workbook,int index) {
        if(cellMapping!=null) {
            CellStyle cellStyle = index % 2 == 0 ? getStyles(false, null) : getStyles(true, null);
            cellStyle = MyCellStyleUtil.onBorder(cellStyle,(short) 1);
            if (cellMapping.getDataType() == BaseEntityTypeConstants.STRING_TYPE) {
                super.createStringCell(row, cellMapping.cellIndex, value == null ? "" : value.toString(),cellStyle,null);
            } else if (cellMapping.getDataType()== BaseEntityTypeConstants.DOUBLE_TYPE) {
                super.createDoubleCell(row, cellMapping.cellIndex, value == null ? "" : value.toString(),cellStyle,null);
            }
            if(mergeLengthways>0 || mergeCrosswise>0) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(index,index+mergeLengthways,cellMapping.cellIndex,cellMapping.cellIndex+mergeCrosswise);
                MyCellStyleUtil.onBorderForMergeCell( 1,cellRangeAddress,sheet,workbook);
                sheet.addMergedRegion(cellRangeAddress);
            }

        }
    }


}
