/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements. See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership. The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.apache.fineract.dataimport.handler;

import org.apache.poi.ss.usermodel.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

/**
 * Created by kyriakos on 4/10/17.
 */
public class WorkbookUtils {
    public static void writeInt(int colIndex, Row row, int value) {
        row.createCell(colIndex).setCellValue(value);
    }

    public static void writeLong(int colIndex, Row row, long value) {
        row.createCell(colIndex).setCellValue(value);
    }

    public static void writeString(int colIndex, Row row, String value) {
        row.createCell(colIndex).setCellValue(value);
    }

    public static void writeDouble(int colIndex, Row row, double value) {
        row.createCell(colIndex).setCellValue(value);
    }

    public static void writeFormula(int colIndex, Row row, String formula) {
        row.createCell(colIndex).setCellType(Cell.CELL_TYPE_FORMULA);
        row.createCell(colIndex).setCellFormula(formula);
    }

    public static CellStyle getDateCellStyle(Workbook workbook) {
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        return dateCellStyle;
    }

    public static void writeDate(int colIndex, Row row, String value, CellStyle dateCellStyle) {
        try {
            //To make validation between functions inclusive.
            Date date = new SimpleDateFormat("dd/MM/yyyy", Locale.ENGLISH).parse(value);
            Calendar cal = Calendar.getInstance();
            cal.setTime(date);
            cal.set(Calendar.HOUR_OF_DAY, 0);
            cal.set(Calendar.MINUTE, 0);
            cal.set(Calendar.SECOND, 0);
            cal.set(Calendar.MILLISECOND, 0);
            Date dateWithoutTime = cal.getTime();
            row.createCell(colIndex).setCellValue(dateWithoutTime);
            row.getCell(colIndex).setCellStyle(dateCellStyle);
        } catch (ParseException pe) {
            throw new IllegalArgumentException("ParseException");
        }
    }

    /*
    public void setClientAndGroupDateLookupTable(Sheet sheet, List<CompactClient> clients, List<CompactGroup> groups,
                                                    int nameCol, int activationDateCol) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        int rowIndex = 0;
        for(CompactClient client: clients) {
            Row row = sheet.getRow(++rowIndex);
            if(row == null)
                row = sheet.createRow(rowIndex);
            writeString(nameCol, row, client.getDisplayName().replaceAll("[ )(] ", "_") + "("+client.getId() + ")");
            writeDate(activationDateCol, row, client.getActivationDate().get(2) + "/" + client.getActivationDate().get(1) + "/" + client.getActivationDate().get(0), dateCellStyle);
        }
        if(groups == null)
            return;
        for(CompactGroup group: groups) {
            Row row = sheet.getRow(++rowIndex);
            if(row == null)
                row = sheet.createRow(rowIndex);
            writeString(nameCol, row, group.getName().replaceAll("[ )(] ", "_"));
            writeDate(activationDateCol, row, group.getActivationDate().get(2) + "/" + group.getActivationDate().get(1) + "/" + group.getActivationDate().get(0), dateCellStyle);
        }
    }

    public void setOfficeDateLookupTable(Sheet sheet, List<Office> offices, int officeNameCol, int activationDateCol) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        int rowIndex = 0;
        for(Office office:offices) {
            Row row = sheet.createRow(++rowIndex);
            writeString(officeNameCol, row, office.getName().trim().replaceAll("[ )(]", "_"));
            writeDate(activationDateCol, row, ""+office.getOpeningDate().get(2)+"/"+office.getOpeningDate().get(1)+"/"+office.getOpeningDate().get(0), dateCellStyle);
        }
    }
    */
}
