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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel

import uk.ac.ox.softeng.maurodatamapper.provider.plugin.AbstractMauroDataMapperPlugin

import java.awt.Color

/**
 * @since 17/08/2017
 */
@SuppressWarnings('SpellCheckingInspection')
class ExcelPlugin extends AbstractMauroDataMapperPlugin {

    public static final String HEADER_COLUMN_COLOUR = '#5B9BD5'
    public static final Color ALTERNATING_COLUMN_COLOUR = Color.decode('#7BABF5')
    public static final Color BORDER_COLOUR = Color.decode('#FFFFFF')
    public static final Double CELL_COLOUR_TINT = 0.6d
    public static final Double BORDER_COLOUR_TINT = -0.35d

    public static final String DATAMODELS_SHEET_NAME = 'DataModels'
    public static final Integer DATAMODELS_HEADER_ROWS = 2
    public static final Integer DATAMODELS_ID_COLUMN = 1

    public static final Integer CONTENT_HEADER_ROWS = 2
    public static final Integer CONTENT_ID_COLUMN = 0

    public static final String DATAMODEL_TEMPLATE_FILENAME = 'Template_DataModel_Import_File.xlsx'

    public static final String DATACLASS_PATH_SPLIT_REGEX = ~/\|/

    @Override
    String getName() {
        'Plugin : Excel'
    }

    @Override
    Closure doWithSpring() {
        { ->
            excelDataModelImporterProviderService(ExcelDataModelImporterProviderService)
            excelDataModelExporterProviderService(ExcelDataModelExporterProviderService)
        }
    }
}
