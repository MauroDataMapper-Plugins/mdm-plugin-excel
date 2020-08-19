package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.rest.controller.ImportParameterGroup
import ox.softeng.metadatacatalogue.core.spi.importer.ImporterService
import ox.softeng.metadatacatalogue.core.type.catalogue.DataModelType
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters

import org.junit.Test

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.assertTrue

/**
 * @since 01/03/2018
 */
@SuppressWarnings('SpellCheckingInspection')
class ExcelImporterServiceTest extends BaseDataModelImportExportExcelTest {

    @Test
    void testImporterParametersAreCorrect() {

        ImporterService importerService = getBean(ImporterService)

        List<ImportParameterGroup> importParameters = importerService.describeImporterParams(importerInstance)

        assertEquals('Number of parameter groups', 2, importParameters.size())

        assertEquals('Number of standard parameters', 3, importParameters[0].size())
        assertEquals('Number of source parameters', 1, importParameters[1].size())

        assertTrue 'Has ImportFile parameter', importParameters[1].first().name == 'importFile'
        assertTrue 'Has Finalised parameter', importParameters[0].any {it.name == 'finalised'}
        assertTrue 'Has importAsNewDocumentationVersion parameter', importParameters[0].any {it.name == 'importAsNewDocumentationVersion'}
        assertTrue 'Has folder parameter', importParameters[0].any {it.name == 'folderId'}
    }

    @Test
    void performSimpleImport() {
        ExcelFileImporterParameters importerParameters = createImportParameters('simpleImport.xlsx')
        DataModel imported = importDomain(importerParameters)

        verifySimpleDataModel DataModel.get(imported.id)
        verifySimpleDataModelContent DataModel.get(imported.id)
    }

    @Test
    void performSimpleImportWithComplexMetadata() {
        ExcelFileImporterParameters importerParameters = createImportParameters('simpleImportComplexMetadata.xlsx')
        DataModel imported = importDomain(importerParameters)

        DataModel dataModel = DataModel.get(imported.id)
        verifySimpleDataModel dataModel, 4
        checkMetadata dataModel, 'ox.softeng.database|dialect', 'test'

        DataClass top = checkDataClass(dataModel, 'top', 3, 'tops description', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.test|blob': 'hello'])
        checkDataElement top, 'info', 'info description', 'string', 1, 1, [
            'different info'    : 'info',
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '20',
        ]
        checkDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        checkDataElement top, 'more info', 'why not', 'string', 0

        DataClass child = checkDataClass(top, 'child', 4, 'child description', 0)
        checkDataElement child, 'info', 'childs info', 'string', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.xsd|min'  : '0',
            'ox.softeng.xsd|max'  : '10',
            'ox.softeng.test|blob': 'wobble'
        ]
        checkDataElement child, 'does it work', 'I don\'t know', 'yesno', 0, -1, [
            'ox.softeng.xsd|choice': 'a',
        ]
        checkDataElement child, 'lazy', 'where the merging only happens in the id col', 'possibly', 1, 1, [
            'different info'       : 'info',
            'ox.softeng.xsd|choice': 'a',
        ]
        checkDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = checkDataClass(top, 'brother', 2)
        checkDataElement brother, 'info', 'brothers info', 'string', 0, 1, [
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '255',]
        checkDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
    }

    @Test
    void performMultipleDataModelImport() {
        ExcelFileImporterParameters importerParameters = createImportParameters('multiDataModelImport.xlsx')
        List<DataModel> importedModels = importDomains(importerParameters, 3)

        assertEquals('Number of DataModels imported', 3, importedModels.size())

        verifySimpleDataModel DataModel.get(importedModels.find {it.label == 'test'}?.id)
        verifySimpleDataModelContent DataModel.get(importedModels.find {it.label == 'test'}?.id)

        DataModel dataModel = DataModel.get(importedModels.find {it.label == 'Another Model'}?.id)

        assertNotNull('Second model must exist', dataModel)

        assertEquals('DataModel name', 'Another Model', dataModel.label)
        assertEquals('DataModel description', 'this is another model which isn\'t overly complicated, and is designed for ' +
                                              'testing dataflows', dataModel.description)
        assertEquals('DataModel author', 'tester', dataModel.author)
        assertEquals('DataModel organisation', 'Oxford', dataModel.organisation)
        assertEquals('DataModel type', DataModelType.DATA_ASSET, dataModel.type)

        assertEquals('DataModel Metadata count', 2, dataModel.metadata.size())

        checkMetadata dataModel, 'reviewed', 'no'
        checkMetadata dataModel, 'distributed', 'no'

        assertEquals('Number of DataTypes', 6, dataModel.dataTypes.size())
        assertEquals('Number of DataClasses', 3, dataModel.dataClasses.size())

        checkDataType dataModel, 'text', 'unlimited text'
        checkDataType dataModel, 'string', 'just a plain old string'
        checkDataType dataModel, 'int', 'a number'
        checkDataType dataModel, 'table', 'a random thing'
        checkEnumeration dataModel, 'yesno', 'an eumeration', [y: 'yes', n: 'no']
        checkDataType dataModel, 'child', null

        DataClass top = checkDataClass(dataModel, 'top', 3, 'tops description', 1, 1, ['extra info': 'some extra info'])
        checkDataElement top, 'info', 'info description', 'string', 1, 1, ['different info': 'info']
        checkDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        checkDataElement top, 'allinfo', 'all the info', 'text'

        DataClass child = checkDataClass(top, 'child', 3, 'child description', 0)
        checkDataElement child, 'info', 'childs info', 'string', 1, 1, ['extra info': 'some extra info']
        checkDataElement child, 'does it work', 'I don\'t know', 'yesno', 1, -1
        checkDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = checkDataClass(top, 'brother', 3)
        checkDataElement brother, 'info', 'brothers info', 'string', 0
        checkDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
        checkDataElement brother, 'twin_sibling', 'reference to the other child', 'child', 0, -1

        dataModel = DataModel.get(importedModels.find {it.label == 'complex.xsd'}?.id)

        assertNotNull('Third model must exist', dataModel)
    }
}
