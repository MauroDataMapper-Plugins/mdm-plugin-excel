/*
 * Copyright 2020-2022 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.core.facet.MetadataService
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClassService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElementService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.enumeration.EnumerationValue
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.security.User
import uk.ac.ox.softeng.maurodatamapper.util.Utils

import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * @since 01/03/2018
 */
@Slf4j
class SimpleExcelDataModelExporterProviderService extends DataModelExporterProviderService {

    DataElementService dataElementService
    DataClassService dataClassService
    MetadataService metadataService

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
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    @Override
    Boolean canExportMultipleDomains() {
        true
    }

    @Override
    String getNamespace() {
        'uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel'
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
            XSSFSheet dataModelsSheet = workbook.createSheet('DataModels')
            XSSFSheet enumerationsSheet = workbook.createSheet('Enumerations')

            List<Map<String, String>> dataModelsSheetArray = []
            List<Map<String, String>> enumerationsSheetArray = []

            dataModels.each {dataModel ->
                String sheetKey = createSheetKey(dataModel.label)
                // Create the sheet for the DM data
                XSSFSheet dataModelSheet = workbook.createSheet(sheetKey)

                // Add row to the DM sheet
                dataModelsSheetArray.add(buildDataModelRow(dataModel, sheetKey))

                // Add the enumerations to the enumeration sheet
                enumerationsSheetArray.addAll(buildEnumerationDataRows(dataModel.label, dataModel.enumerationTypes))

                List<Map<String, String>> dataModelSheetArray = []

                dataModel.childDataClasses.each {dataClass ->
                    dataModelSheetArray.addAll(buildDataClassRows(dataClass, null))
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

    Map<String, String> buildDataModelRow(DataModel dataModel, String sheetKey) {
        Map<String, String> dataRow = [
            Name        : dataModel.label,
            Description : dataModel.description,
            Author      : dataModel.author,
            Organisation: dataModel.organisation,
            'Sheet Key' : sheetKey,
            Type        : dataModel.modelType,
        ]

        metadataService.findAllByMultiFacetAwareItemId(dataModel.id).each {metadata ->
            String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }
        dataRow
    }

    List<Map<String, String>> buildEnumerationDataRows(String dataModelLabel, Collection<EnumerationType> enumerationTypes) {
        enumerationTypes.collect {enumerationType ->
            enumerationType.enumerationValues.collect {enumValue ->
                buildEnumerationValueRow(dataModelLabel, enumerationType, enumValue)
            }
        }.flatten() as List<Map<String, String>>
    }

    Map<String, String> buildEnumerationValueRow(String dataModelLabel, EnumerationType enumerationType, EnumerationValue enumerationValue) {
        Map<String, String> dataRow = [
            'DataModel Name'  : dataModelLabel,
            'Enumeration Name': enumerationType.label,
            Description       : enumerationType.description,
            Key               : enumerationValue.key,
            Value             : enumerationValue.value,
        ]
        metadataService.findAllByMultiFacetAwareItemId(enumerationValue.id).each {metadata ->
            String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }
        dataRow
    }

    List<Map<String, String>> buildDataClassRows(DataClass dataClass, String path) {

        List<Map<String, String>> dataClassSheetArray = []

        String dataClassPath = (!path) ? dataClass.label : ("${path} | ${dataClass.label}")

        Map<String, String> dataRow = [
            'DataClass Path'          : dataClassPath,
            'DataElement Name'        : '',
            Description               : dataClass.description,
            'Minimum \r\nMultiplicity': dataClass.minMultiplicity || dataClass.minMultiplicity == 0 ? dataClass.minMultiplicity.toString() : '',
            'Maximum \r\nMultiplicity': dataClass.maxMultiplicity || dataClass.maxMultiplicity == 0 ? dataClass.maxMultiplicity.toString() : '',
            'DataType Name'           : '',
            'DataType Reference'      : ''
        ]

        metadataService.findAllByMultiFacetAwareItemId(dataClass.id).each {metadata ->
            String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }

        dataClassSheetArray.add(dataRow)

        dataClass.dataElements?.each {dataElement ->
            Map dataElementDataRow = [
                'DataClass Path'          : dataClassPath,
                'DataElement Name'        : dataElement.label,
                Description               : dataElement.description,
                'Minimum \r\nMultiplicity': dataElement.minMultiplicity || dataElement.minMultiplicity == 0 ? dataElement.minMultiplicity.toString() : '',
                'Maximum \r\nMultiplicity': dataElement.maxMultiplicity || dataElement.maxMultiplicity == 0 ? dataElement.maxMultiplicity.toString() : '',
                'DataType Name'           : dataElement.dataType.label,
            ]
            if (dataElement.dataType instanceof ReferenceType) {
                List<String> classPath = getClassPath(((ReferenceType) dataElement.dataType).referenceClass, [])
                dataElementDataRow['DataType Reference'] = classPath.join(' | ')
            }

            metadataService.findAllByMultiFacetAwareItemId(dataElement.id).each {metadata ->
                String key = "${metadata.namespace}:${metadata.key}"
                dataElementDataRow[key] = metadata.value
            }
            dataClassSheetArray.add(dataElementDataRow)
        }
        dataClass.dataClasses?.each {childDataClass ->
            dataClassSheetArray.addAll(buildDataClassRows(childDataClass, dataClassPath))
        }
        dataClassSheetArray
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

    void writeArrayToSheet(Workbook workbook, XSSFSheet sheet, List<Map<String, String>> array) {
        long start = System.currentTimeMillis()
        Set<String> headers = [] as Set
        if (array.size() > 0) {
            headers.addAll(array.get(0).keySet())
        }
        array.eachWithIndex {entry, i ->
            if (i != 0) {
                headers.addAll(entry.keySet())
            }
        }
        Row headerRow = sheet.createRow(0)
        headers.eachWithIndex {header, idx ->
            Cell headerCell = headerRow.createCell(idx)
            headerCell.setCellValue(header)
        }
        setHeaderRow(headerRow, workbook)

        long substart = System.currentTimeMillis()

        array.eachWithIndex {map, rowIdx ->

            Row valueRow = sheet.createRow(rowIdx + 1)
            headers.eachWithIndex {header, cellIdx ->
                Cell valueCell = valueRow.createCell(cellIdx)
                valueCell.setCellValue(map[header] ?: '')
            }
        }
        log.debug('Writing rows took {}', Utils.timeTaken(substart))
        // Disabled this as it adds 1/3 of the time to run
        // substart = System.currentTimeMillis()

        // autoSizeColumns(sheet, 0..(headers.size() + 1))
        // log.debug('Resizing columns took {}',  Utils.timeTaken(substart))
        log.debug('Writing array of {} rows to sheet took {}', array.size(), Utils.timeTaken(start))
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
        columns.each {sheet.autoSizeColumn(it, true)}
    }

    static String createSheetKey(String dataModelName) {

        String sheetKey = ''
        String[] words = dataModelName.split(' ')
        if (words.size() == 1) {
            sheetKey = dataModelName.toUpperCase()
        } else {
            words.each {word ->
                if (word.length() > 0) {
                    String newWord = word.replaceAll('^[^A-Za-z]*', '')

                    sheetKey += newWord[0].toUpperCase() + newWord.replaceAll('[^0-9]', '')
                }
            }
        }
        return sheetKey
    }
}
