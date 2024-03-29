/*
 * Copyright 2020-2023 University of Oxford and NHS England
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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer

import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.BaseExcelDataModelImporterExporterProviderServiceSpec
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter.SimpleExcelDataModelExporterProviderService

import grails.gorm.transactions.Rollback
import grails.testing.mixin.integration.Integration
import groovy.util.logging.Slf4j

@Slf4j
@Integration
@Rollback
class SimpleExcelDataModelImporterProviderServiceSpec extends BaseExcelDataModelImporterExporterProviderServiceSpec {

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

    boolean isSimpleImport(){
        true
    }

    def 'testSimpleImport'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel = importAndValidateModel(createImportParameters('simpleImport.simple.xlsx'))
        then:
        verifySimpleDataModel dataModel
        verifySimpleDataModelContent dataModel
    }

    def 'testSimpleImportWithComplexMetadata'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel =importAndValidateModel(createImportParameters('simpleImportComplexMetadata.simple.xlsx'))
        then:
        verifySimpleDataModelWithComplexMetadata dataModel
        verifySimpleDataModelWithComplexMetadataContent dataModel
    }

    def 'testMultipleDataModelImport'() {
        given:
        setupDomainData()

        when:
        List<DataModel> dataModels = importAndValidateModels(createImportParameters('multiDataModelImport.simple.xlsx'))

        then:
        dataModels.size() == 3

        when:
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