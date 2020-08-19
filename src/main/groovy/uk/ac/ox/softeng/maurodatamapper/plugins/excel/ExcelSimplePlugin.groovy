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

import ox.softeng.metadatacatalogue.core.spi.module.AbstractModule

/**
 * @since 17/08/2017
 */
@SuppressWarnings('SpellCheckingInspection')
class ExcelSimplePlugin extends AbstractModule {

    @Override
    String getName() {
        'Plugin : Simple Excel'
    }

    @Override
    Closure doWithSpring() {
        {->
            excelSimpleImporterService(ExcelSimpleDataModelImporterService)
            excelSimpleExporterService(ExcelSimpleDataModelExporterService)

        }
    }

    static DATAMODEL_SHEET_COLUMNS = ["Name", "Description", "Author", "Organisation", "Sheet Key", "Type"]
    static ENUM_SHEET_COLUMNS = ["DataModel Name", "Name", "Description", "Key", "Value"]
    static MODEL_SHEET_COLUMNS = ["DataClass Path", "Name", "Description", "Minimum Multiplicity", "Maximum Multiplicity", "DataType Name",
                                  "DataType Reference"]

}
