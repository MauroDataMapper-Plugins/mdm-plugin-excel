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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter


import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.BaseExcelDataModelImporterExporterProviderServiceSpec
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer.SimpleExcelDataModelImporterProviderService

import grails.gorm.transactions.Rollback
import grails.testing.mixin.integration.Integration
import groovy.util.logging.Slf4j

@Slf4j
@Integration
@Rollback
class SimpleExcelDataModelExporterProviderServiceSpec extends BaseExcelDataModelImporterExporterProviderServiceSpec {

    private static final String EXPORT_FILEPATH = 'build/tmp/'

    SimpleExcelDataModelExporterProviderService simpleExcelDataModelExporterProviderService
    SimpleExcelDataModelImporterProviderService simpleExcelDataModelImporterProviderService

    @Override
    DataModelImporterProviderService getDataModelImporterProviderService() {
        simpleExcelDataModelImporterProviderService
    }

    @Override
    DataModelExporterProviderService getDataModelExporterProviderService() {
        simpleExcelDataModelExporterProviderService
    }

    String getMetadataKeySplit(){
        '[:|]'
    }

    boolean isSimpleImport(){
        true
    }

    def 'testSimpleExport'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel = importThenExportWorkbook('simpleImport.simple.xlsx', 'simpleImport_export.simple.xlsx') as DataModel

        then:
        verifySimpleDataModel dataModel
        verifySimpleDataModelContent dataModel
    }

    def 'testSimpleExportWithComplexMetadata'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel = importThenExportWorkbook(
            'simpleImportComplexMetadata.simple.xlsx', 'simpleImportComplexMetadata_export.simple.xlsx') as DataModel

        then:
        verifySimpleDataModelWithComplexMetadata dataModel
        verifySimpleDataModelWithComplexMetadataContent dataModel
    }

    def 'testMultipleDataModelExport'() {
        given:
        setupDomainData()

        when:
        List<DataModel> dataModels = importThenExportWorkbook('multiDataModelImport.simple.xlsx', 'multiDataModelImport_export.simple.xlsx') as List<DataModel>

        DataModel simpleDataModel = findByLabel(dataModels, 'test')

        then:
        verifySimpleDataModel simpleDataModel
        verifySimpleDataModelContent simpleDataModel

        when:
        DataModel dataFlowDataModel = findByLabel(dataModels, 'Another Model')

        then:
        verifyDataFlowDataModel dataFlowDataModel
        verifyDataFlowDataModelContent dataFlowDataModel

        when:
        DataModel complexDataModel = findByLabel(dataModels, 'complex.xsd')

        then:
        verifyComplexDataModel complexDataModel
        verifyComplexDataModelContent complexDataModel
    }
}
