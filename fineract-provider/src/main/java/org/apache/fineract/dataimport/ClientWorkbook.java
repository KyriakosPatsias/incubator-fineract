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
package org.apache.fineract.dataimport;

import org.apache.fineract.dataimport.handler.WorkbookUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import org.apache.fineract.dataimport.handler.ClientSheetPopulatorHandler;

/**
 * Created by kyriakos on 4/10/17.
 */

@Service
public class ClientWorkbook {

    private String clientType;

    private Workbook workbook;

    private static final int FIRST_NAME_COL = 0;
    private static final int LAST_NAME_COL = 1;
    private static final int MIDDLE_NAME_COL = 2;
    private static final int FULL_NAME_COL = 0;
    private static final int OFFICE_NAME_COL = 3;
    private static final int STAFF_NAME_COL = 4;
    private static final int EXTERNAL_ID_COL = 5;
    private static final int ACTIVATION_DATE_COL = 6;
    private static final int ACTIVE_COL = 7;
    private static final int WARNING_COL = 9;
    private static final int RELATIONAL_OFFICE_NAME_COL = 16;
    private static final int RELATIONAL_OFFICE_OPENING_DATE_COL = 17;

    private ClientSheetPopulatorHandler csph;

    @Autowired
    public ClientWorkbook(){

        workbook = new HSSFWorkbook();
        Sheet clientSheet = workbook.createSheet("Clients");
        setLayout(clientSheet);
    }

    public void setClientType(String clientType) {
        this.clientType = clientType;
    }

    public Workbook getTemplate() {
        return workbook;
    }

    private void setLayout(Sheet worksheet) {
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        if(clientType.equals("individual")) {
            worksheet.setColumnWidth(FIRST_NAME_COL, 6000);
            worksheet.setColumnWidth(LAST_NAME_COL, 6000);
            worksheet.setColumnWidth(MIDDLE_NAME_COL, 6000);
            WorkbookUtils.writeString(FIRST_NAME_COL, rowHeader, "First Name*");
            WorkbookUtils.writeString(LAST_NAME_COL, rowHeader, "Last Name*");
            WorkbookUtils.writeString(MIDDLE_NAME_COL, rowHeader, "Middle Name");
        }
        else {
            worksheet.setColumnWidth(FULL_NAME_COL, 10000);
            worksheet.setColumnWidth(LAST_NAME_COL, 0);
            worksheet.setColumnWidth(MIDDLE_NAME_COL, 0);
            WorkbookUtils.writeString(FULL_NAME_COL, rowHeader, "Full/Business Name*");
        }
        worksheet.setColumnWidth(OFFICE_NAME_COL, 5000);
        worksheet.setColumnWidth(STAFF_NAME_COL, 5000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL, 3500);
        worksheet.setColumnWidth(ACTIVATION_DATE_COL, 4000);
        worksheet.setColumnWidth(ACTIVE_COL, 2000);
        worksheet.setColumnWidth(RELATIONAL_OFFICE_NAME_COL, 6000);
        worksheet.setColumnWidth(RELATIONAL_OFFICE_OPENING_DATE_COL, 4000);
        WorkbookUtils.writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        WorkbookUtils.writeString(STAFF_NAME_COL, rowHeader, "Staff Name*");
        WorkbookUtils.writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
        WorkbookUtils.writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
        WorkbookUtils.writeString(ACTIVE_COL, rowHeader, "Active*");
        WorkbookUtils.writeString(WARNING_COL, rowHeader, "All * marked fields are compulsory.");
        WorkbookUtils.writeString(RELATIONAL_OFFICE_NAME_COL, rowHeader, "Office Name");
        WorkbookUtils.writeString(RELATIONAL_OFFICE_OPENING_DATE_COL, rowHeader, "Opening Date");

    }
}
