/*
 * Copyright 2020 University of Oxford
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * SPDX-License-Identifier: Apache-2.0
 */
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.simple

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClassService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElementService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceType
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.beans.factory.annotation.Autowired

import groovy.util.logging.Slf4j

/**
 * @since 01/03/2018
 */
@Slf4j
class ExcelSimpleDataModelExporterProviderService extends DataModelExporterProviderService {

    @Autowired
    DataModelService dataModelService

    @Autowired
    DataElementService dataElementService

    @Autowired
    DataClassService dataClassService

    @Override
    String getDisplayName() {
        'Simple Excel (XLSX) Exporter'
    }

    @Override
    String getFileExtension() {
        'xlsx'
    }

    @Override
    String getFileType() {
        'application/vnd.ms-excel'
    }

    @Override
    String getVersion() {
        '1.0.0'
    }

    @Override
    Boolean canExportMultipleDomains() {
        true
    }

    @Override
    ByteArrayOutputStream exportDataModel(User currentUser, DataModel dataModel) throws ApiException {
        exportDataModels(currentUser, [dataModel])
    }

    @Override
    ByteArrayOutputStream exportDataModels(User currentUser, List<DataModel> dataModels) throws ApiException {

        log.info('Exporting DataModels to Excel')
        XSSFWorkbook workbook = null
        try {
            workbook = new XSSFWorkbook()
            XSSFSheet dataModelsSheet = workbook.createSheet("DataModels")
            XSSFSheet enumerationsSheet = workbook.createSheet("Enumerations")

            List<LinkedHashMap<String, String>> dataModelsSheetArray = []
            List<LinkedHashMap<String, String>> enumerationsSheetArray = []

            dataModels.each { dataModel ->
                String sheetKey = createSheetKey(dataModel.label)
                LinkedHashMap<String, String> array = new LinkedHashMap<String, String>()
                array["Name"] = dataModel.label
                array["Description"] = dataModel.description
                array["Author"] = dataModel.author
                array["Organisation"] = dataModel.organisation
                array["Sheet Key"] = sheetKey
                array["Type"] = dataModel.type.toString()

                dataModel.metadata.each { metadata ->
                    String key = "${metadata.namespace}:${metadata.key}"
                    array[key] = metadata.value
                }
                dataModelsSheetArray.add(array)

                List<LinkedHashMap<String, String>> enumerationsArray = []
                dataModel.dataTypes.findAll { it instanceof EnumerationType }.each { dataType ->
                    EnumerationType enumType = (EnumerationType) dataType
                    enumType.enumerationValues.each { enumValue ->
                        LinkedHashMap<String, String> enumArray = new LinkedHashMap<String, String>()
                        enumArray["DataModel Name"] = dataModel.label
                        enumArray["Name"] = enumType.label
                        enumArray["Description"] = enumType.description
                        enumArray["Key"] = enumValue.key
                        enumArray["Value"] = enumValue.value
                        enumValue.metadata.each { metadata ->
                            String key = "${metadata.namespace}:${metadata.key}"
                            enumArray[key] = metadata.value
                        }
                        enumerationsArray.add(enumArray)
                    }
                }
                enumerationsSheetArray.addAll(enumerationsArray)

                XSSFSheet dataModelSheet = workbook.createSheet(sheetKey)
                List<LinkedHashMap<String, String>> dataModelSheetArray = []

                dataModel.childDataClasses.each { dataClass ->
                    dataModelSheetArray.addAll(createArrayFromClass(dataClass, null))
                }
                writeArrayToSheet(workbook, dataModelSheet, dataModelSheetArray)
            }

            writeArrayToSheet(workbook, dataModelsSheet, dataModelsSheetArray)
            writeArrayToSheet(workbook, enumerationsSheet, enumerationsSheetArray)

            if (workbook) {
                ByteArrayOutputStream os = new ByteArrayOutputStream(1000000)
                workbook.write(os)
                log.info('DataModels exported')
                return os
            }
        } finally {
            closeWorkbook(workbook)
        }
        null
    }

    void closeWorkbook(Workbook workbook) {
        if (workbook != null) {
            try {
                workbook.close()
            } catch (IOException ignored) {
                // ignored
            }
        }
    }

    void autoSizeColumns(Sheet sheet, List<Integer> columns) {
        columns.each { sheet.autoSizeColumn(it, true) }
    }

    void writeArrayToSheet(Workbook workbook, XSSFSheet sheet, List<LinkedHashMap<String, String>> array) {
        LinkedHashSet<String> headers = new LinkedHashSet<String>()
        if (array.size() > 0) {
            headers.addAll(array.get(0).keySet())
        }
        array.eachWithIndex { entry, i ->
            if (i != 0) {
                headers.addAll(entry.keySet())
            }
        }
        Row headerRow = sheet.createRow(0)
        headers.eachWithIndex { header, idx ->
            Cell headerCell = headerRow.createCell(idx)
            headerCell.setCellValue(header)
        }
        setHeaderRow(headerRow, workbook)

        array.eachWithIndex { map, rowIdx ->
            Row valueRow = sheet.createRow(rowIdx + 1)
            headers.eachWithIndex { header, cellIdx ->
                Cell valueCell = valueRow.createCell(cellIdx)
                valueCell.setCellValue(map[header] ?: "")
            }
        }
        autoSizeColumns(sheet, 0..(headers.size() + 1))
    }

    static String createSheetKey(String dataModelName) {

        String sheetKey = ''
        String[] words = dataModelName.split(' ')
        if (words.size() == 1) {
            sheetKey = dataModelName.toUpperCase()
        } else {
            words.each { word ->
                if (word.length() > 0) {
                    String newWord = word.replaceAll("^[^A-Za-z]*", "")

                    sheetKey += newWord[0].toUpperCase() + newWord.replaceAll("[^0-9]", "")
                }
            }
        }
        return sheetKey
    }

    List<LinkedHashMap<String, String>> createArrayFromClass(DataClass dc, String path) {

        List<LinkedHashMap<String, String>> dataClassSheetArray = []

        String dataClassPath = (!path) ? dc.label : (path + " | " + dc.label)

        LinkedHashMap<String, String> array = new LinkedHashMap<String, String>()

        array["DataClass Path"] = dataClassPath
        array["Name"] = "" // dc.label
        array["Description"] = dc.description
        if (dc.minMultiplicity || dc.minMultiplicity == 0) {
            array["Minimum Multiplicity"] = dc.minMultiplicity.toString()
        } else {
            // Empty assignment helps ensure column order
            array["Minimum Multiplicity"] = ""
        }
        if (dc.maxMultiplicity || dc.minMultiplicity == 0) {
            array["Maximum Multiplicity"] = dc.maxMultiplicity.toString()
        } else {
            // Empty assignment helps ensure column order
            array["Maximum Multiplicity"] = ""
        }

        dc.metadata.each { metadata ->
            String key = "${metadata.namespace}:${metadata.key}"
            array[key] = metadata.value
        }

        dataClassSheetArray.add(array)

        dc.childDataElements?.each { dataElement ->
            array = new LinkedHashMap<String, String>()

            array["DataClass Path"] = dataClassPath
            array["Name"] = dataElement.label
            array["Description"] = dataElement.description
            if (dataElement.minMultiplicity || dataElement.minMultiplicity == 0) {
                array["Minimum Multiplicity"] = dataElement.minMultiplicity.toString()
            }
            if (dataElement.maxMultiplicity || dataElement.maxMultiplicity == 0) {
                array["Maximum Multiplicity"] = dataElement.maxMultiplicity.toString()
            }
            array["DataType Name"] = dataElement.dataType.label
            if (dataElement.dataType instanceof ReferenceType) {
                List<String> classPath = getClassPath(((ReferenceType) dataElement.dataType).referenceClass, [])
                array["DataType Reference"] = StringUtils.join(classPath, " | ")
            }

            dataElement.metadata.each { metadata ->
                String key = "${metadata.namespace}:${metadata.key}"
                array[key] = metadata.value
            }
            dataClassSheetArray.add(array)
        }
        dc.childDataClasses.each { childDataClass ->
            dataClassSheetArray.addAll(createArrayFromClass(childDataClass, dataClassPath))
        }
        return dataClassSheetArray
    }

    void setHeaderRow(Row row, Workbook workbook) {

        CellStyle style = row.getRowStyle()
        if (!style) {
            style = workbook.createCellStyle()
            row.setRowStyle(style)
        }

        Font font = workbook.createFont()
        font.setBold(true)
        style.setFont(font)
        style.setFillForegroundColor((short) 123456)
        row.setRowStyle(style)

        for (int i = 0; i < row.getLastCellNum(); i++) {
            row.getCell(i).setCellStyle(style)
        }
    }

    List<String> getClassPath(DataClass dataClass, List<String> path) {
        path.add(0, dataClass.label)
        if (dataClass.parentDataClass) {
            return getClassPath(dataClass.parentDataClass, path)
        } else {
            return path
        }
    }
}
