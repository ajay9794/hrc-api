package com.herambco.hrcapi.service.implementation;

import com.herambco.hrcapi.Dto.XlsFilterResponseDto;
import com.herambco.hrcapi.service.interfaces.FilterXLSInterface;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.atomic.AtomicReference;

@Service
public class FilterXLSImpl implements FilterXLSInterface {
    @Override
    public List<Object> generateXLSBasedOnFilter(MultipartFile file, String itemFilter, Boolean gstFilter) {
        Workbook workbook = getWorkBook(file);
        Sheet sheet = workbook.getSheetAt(0);
        List<Object> responseList = new ArrayList<>();
        AtomicReference<Boolean> rowFound = new AtomicReference<>(false);
        try{
            if(gstFilter) {
                List<XlsFilterResponseDto> twoPointFiveGST = new ArrayList<>();
                List<XlsFilterResponseDto> sixGST = new ArrayList<>();
                List<XlsFilterResponseDto> sevenGST = new ArrayList<>();
                List<XlsFilterResponseDto> nineGST = new ArrayList<>();
                List<XlsFilterResponseDto> blankGST = new ArrayList<>();
                sheet.forEach((Row row) -> {
                    if(rowFound.get()) {
                        String gstPercentage = "";
                        if(row.getCell(9) != null) {
                            gstPercentage = row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue()) : row.getCell(9).getStringCellValue();
                        }


                        switch(gstPercentage) {
                            case "2.5":
                                twoPointFiveGST.add(new XlsFilterResponseDto(
                                        row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                        row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                        row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue() * 2) : NumberToTextConverter.toText(Double.parseDouble(row.getCell(9).getStringCellValue()) * 2),
                                        row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                        row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                                break;
                            case "6":
                                sixGST.add(new XlsFilterResponseDto(
                                        row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                        row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                        row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue() * 2) : NumberToTextConverter.toText(Double.parseDouble(row.getCell(9).getStringCellValue()) * 2),
                                        row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                        row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                                break;
                            case "7":
                                sevenGST.add(new XlsFilterResponseDto(
                                        row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                        row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                        row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue() * 2) : NumberToTextConverter.toText(Double.parseDouble(row.getCell(9).getStringCellValue()) * 2),
                                        row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                        row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                                break;
                            case "9":
                                nineGST.add(new XlsFilterResponseDto(
                                        row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                        row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                        row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue() * 2) : NumberToTextConverter.toText(Double.parseDouble(row.getCell(9).getStringCellValue()) * 2),
                                        row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                        row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                                break;
                            default:
                                blankGST.add(new XlsFilterResponseDto(
                                        row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                        row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                        row.getCell(9) == null ? "" : row.getCell(9).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(9).getNumericCellValue()) : row.getCell(9).getStringCellValue(),
                                        row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                        row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                                break;
                        }
                    } else if(row.getCell(4) != null) {
                        if (row.getCell(4).getCellType() == CellType.NUMERIC && NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()).contains("Item Name")) {
                            rowFound.set(true);
                        } else if (row.getCell(4).getCellType() == CellType.STRING && row.getCell(4).getStringCellValue().contains("Item Name")) {
                            rowFound.set(true);
                        }
                    }
                });
                responseList.add(twoPointFiveGST);
                responseList.add(sixGST);
                responseList.add(sevenGST);
                responseList.add(nineGST);
                responseList.add(blankGST);
            } else if(StringUtils.isNotEmpty(itemFilter)) {
                List<XlsFilterResponseDto> itemNameFilterList = new ArrayList<>();
                sheet.forEach((Row row) -> {
                    if(rowFound.get()) {
                        String itemName =  row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue();
                        if(StringUtils.isNotEmpty(itemName) && itemName.equalsIgnoreCase(itemFilter)) {
                            itemNameFilterList.add(new XlsFilterResponseDto(
                                    row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue(),
                                    row.getCell(6) == null ? "" : row.getCell(6).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()) : row.getCell(6).getStringCellValue(),
                                    "",
                                    row.getCell(7) == null ? null : row.getCell(7).getCellType() == CellType.NUMERIC ? row.getCell(7).getNumericCellValue() : Double.parseDouble(row.getCell(7).getStringCellValue()),
                                    row.getCell(14) == null ? "" : row.getCell(14).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(14).getNumericCellValue()) : row.getCell(14).getStringCellValue()));
                        }
                    }  else if(row.getCell(4) != null) {
                        if (row.getCell(4).getCellType() == CellType.NUMERIC && NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()).contains("Item Name")) {
                            rowFound.set(true);
                        } else if (row.getCell(4).getCellType() == CellType.STRING && row.getCell(4).getStringCellValue().contains("Item Name")) {
                            rowFound.set(true);
                        }
                    }
                });
                responseList.add(itemNameFilterList);
            }

        } catch(Exception e) {
            e.printStackTrace();
        }

        return responseList;
    }

    @Override
    public Set<String> getItemList(MultipartFile file) {
        Workbook workbook = getWorkBook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Set<String> responseList = new HashSet<>();
        try{
            AtomicReference<Boolean> rowFound = new AtomicReference<>(false);
            sheet.forEach((Row row) -> {
                if(rowFound.get()) {
                    responseList.add(
                            row.getCell(4) == null ? "" : row.getCell(4).getCellType() == CellType.NUMERIC ? NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()) : row.getCell(4).getStringCellValue()
                    );
                } else if(row.getCell(4) != null) {
                    if (row.getCell(4).getCellType() == CellType.NUMERIC && NumberToTextConverter.toText(row.getCell(4).getNumericCellValue()).contains("Item Name")) {
                        rowFound.set(true);
                    } else if (row.getCell(4).getCellType() == CellType.STRING && row.getCell(4).getStringCellValue().contains("Item Name")) {
                        rowFound.set(true);
                    }
                }
            });
        }catch(Exception e) {
            e.printStackTrace();
        }
        return responseList;
    }

    private Workbook getWorkBook(MultipartFile file) {
        Workbook workbook = null;
        String extension = FilenameUtils.getExtension(file.getOriginalFilename());
        try {
            assert extension != null;
            if (extension.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(file.getInputStream());
            } else if (extension.equalsIgnoreCase("xls")) {
                workbook = new HSSFWorkbook(file.getInputStream());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }
}
