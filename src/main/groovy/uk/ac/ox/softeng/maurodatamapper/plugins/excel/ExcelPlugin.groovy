/*
 * Copyright 2020-2021 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.simple.ExcelSimpleDataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.simple.ExcelSimpleDataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.provider.plugin.AbstractMauroDataMapperPlugin

import groovy.transform.CompileDynamic
import groovy.transform.CompileStatic

@CompileStatic
@SuppressWarnings('SpellCheckingInspection')
class ExcelPlugin extends AbstractMauroDataMapperPlugin {

    static final String CHARSET = 'ISO-8859-1'
    static final String EXCEL_FILETYPE = 'application/vnd.ms-excel'
    static final String EXCEL_FILE_EXTENSION = 'xlsx'

    static final String DATAMODELS_IMPORT_TEMPLATE_FILENAME = 'Template_DataModel_Import_File.xlsx'
    static final String DATAMODELS_SHEET_NAME = 'DataModels'
    static final String CONTENT_TEMPLATE_SHEET_NAME = 'KEY_1'

    static final int DATAMODELS_NUM_HEADER_ROWS = 2
    static final int CONTENT_NUM_HEADER_ROWS = 2
    static final int DATAMODELS_ID_COLUMN_INDEX = 1
    static final int CONTENT_ID_COLUMN_INDEX = 0

    @Override
    String getName() {
        'Plugin : Excel'
    }

    @Override
    @CompileDynamic
    Closure doWithSpring() {
        { ->
            excelDataModelImporterProviderService ExcelDataModelImporterProviderService
            excelDataModelExporterProviderService ExcelDataModelExporterProviderService
            excelSimpleDataModelImporterProviderService ExcelSimpleDataModelImporterProviderService
            excelSimpleDataModelExporterProviderService ExcelSimpleDataModelExporterProviderService
        }
    }
}
