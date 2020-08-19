package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponent
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.rest.controller.ImportParameterGroup
import ox.softeng.metadatacatalogue.core.spi.importer.ImporterService
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.DataFlowImporterPluginParameters
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelDataFlowImporterParameters
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters

import org.hibernate.SessionFactory
import org.junit.Test

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.assertTrue
import static org.junit.Assert.fail

/**
 * @since 09/03/2018
 */
class ExcelDataFlowImporterServiceTest extends BaseImportExportExcelTest<DataFlow, ExcelDataFlowImporterParameters, ExcelDataFlowImporterService> {

    @Test
    void testImporterParametersAreCorrect() {

        ImporterService importerService = getBean(ImporterService)

        List<ImportParameterGroup> importParameters = importerService.describeImporterParams(importerInstance)

        assertEquals('Number of parameters', 1, importParameters.size())

        assertTrue 'Has ImportFile parameter', importParameters[0].first().name == 'importFile'
    }

    @Test
    void performSimpleImport() {

        importTestDataModels()

        DataFlowImporterPluginParameters importerParameters = createImportParameters('Simple_DataFlow_Import_File.xlsx')
        DataFlow imported = importDomain(importerParameters)

        DataFlow dataFlow = DataFlow.get(imported.id)

        verifySimpleDataFlow(dataFlow)
    }

    @Test
    void performSimpleImportAndUpdate() {

        importTestDataModels()

        DataFlowImporterPluginParameters importerParameters = createImportParameters('Simple_DataFlow_Import_File.xlsx')

        importDomain(importerParameters)

        assertEquals("Only 1 DataFlow with the name", 1, DataFlow.countByLabel('Sample DataFlow'))

        importDomain(importerParameters as ExcelDataFlowImporterParameters)
        assertEquals("Only 1 DataFlow with the name", 1, DataFlow.countByLabel('Sample DataFlow'))
    }

    @Test
    void performMultiDataFlowImport() {
        importTestDataModels()

        DataFlowImporterPluginParameters importerParameters = createImportParameters('Multiple_DataFlow_Import_File.xlsx')
        List<DataFlow> imported = importDomains(importerParameters, 2)

        DataFlow dataFlow = DataFlow.get(imported.find {it.label == 'Sample DataFlow'}?.id)

        verifySimpleDataFlow(dataFlow)

        dataFlow = DataFlow.get(imported.find {it.label == 'Second DataFlow'}?.id)

        assertNotNull('Second model must exist', dataFlow)

        assertEquals('DataFlow name', 'Second DataFlow', dataFlow.label)
        assertEquals('DataFlow description', 'This is a second dataflow between the another model and complex model',
                     dataFlow.description)

        assertEquals('DataFlow source', 'complex.xsd', dataFlow.source.label)
        assertEquals('DataFlow target', 'Another Model', dataFlow.target.label)

        assertEquals("Number of DataFlow Components", 9, dataFlow.dataFlowComponents.size())
    }

    protected void verifySimpleDataFlow(DataFlow dataFlow) {
        assertNotNull('Simple model must exist', dataFlow)
        assertEquals('DataFlow name', 'Sample DataFlow', dataFlow.label)
        assertEquals('DataFlow description', 'This is a sample dataflow which shows all the possible ways we can show data',
                     dataFlow.description)

        assertEquals('DataFlow source', 'test', dataFlow.source.label)
        assertEquals('DataFlow target', 'Another Model', dataFlow.target.label)

        assertEquals("Number of DataFlow Components", 8, dataFlow.dataFlowComponents.size())

        Collection<DataFlowComponent> dfcs = dataFlow.dataFlowComponents.findAll {it.label == 'merge'}

        assertEquals("Number of merge transforms", 2, dfcs.size())

        assertTrue("Merge with 3 source & 1 target", dfcs.any {
            it.sourceElements.size() == 3 && it.targetElements.size() == 1 &&
            it.description == 'merge all element info to one location, we will be keeping the elements but we want all the info together as ' +
            'well'
        })
        assertTrue("Merge with 2 source & 1 target", dfcs.any {
            it.sourceElements.size() == 2 && it.targetElements.size() == 1 &&
            it.description == 'merge 2 info elements into one' &&
            it.metadata.size() == 1 &&
            it.metadata.first().key == 'merge separator' &&
            it.metadata.first().value == ','
        })

        dfcs = dataFlow.dataFlowComponents.findAll {it.label == 'copy'}

        assertEquals("Number of copy transforms", 3, dfcs.size())

        assertTrue("Copy another", dfcs.any {it.sourceElements.first().label == 'another' && it.targetElements.first().label == 'another'})
        assertTrue("Copy info child", dfcs.any {
            it.sourceElements.first().label == 'info' &&
            it.targetElements.first().label == 'info' &&
            it.sourceElements.first().dataClass.label == 'child' &&
            it.targetElements.first().dataClass.label == 'child'
        })
        assertTrue("Copy info brother", dfcs.any {
            it.sourceElements.first().label == 'info' &&
            it.targetElements.first().label == 'info' &&
            it.sourceElements.first().dataClass.label == 'brother' &&
            it.targetElements.first().dataClass.label == 'brother'
        })
        assertTrue("Copy description", dfcs.every {it.description == 'copy with no changes'})

        DataFlowComponent dfc = dataFlow.dataFlowComponents.find {it.label == 'required field'}

        assertNotNull("Required field should exist", dfc)
        assertEquals('required field description', 'if empty then set to the default', dfc.description)
        assertEquals('required field source', 'does it work', dfc.sourceElements.first().label)
        assertEquals('required field target', 'does it work', dfc.targetElements.first().label)
        checkMetadata(dfc, 'default', 'y')

        dfc = dataFlow.dataFlowComponents.find {it.label == 'clean'}

        assertNotNull("clean should exist", dfc)
        assertEquals('clean description', 'this field contains dangerous data which needs to be cleaned', dfc.description)
        assertEquals('clean source', 'wrong order', dfc.sourceElements.first().label)
        assertEquals('clean target', 'wrong order', dfc.targetElements.first().label)
        checkMetadata(dfc, 'cleaning technique', 'all identifiable data')

        dfc = dataFlow.dataFlowComponents.find {it.label == 'duplicate'}

        assertNotNull("duplicate should exist", dfc)
        assertEquals('duplicate description', 'we need to duplicate this data into another field', dfc.description)
        assertEquals('duplicate source', 'sibling', dfc.sourceElements.first().label)
        assertEquals('duplicate targets size', 2, dfc.targetElements.size())
        assertTrue('duplicate target 1', dfc.targetElements.any {it.label == 'sibling' && it.dataClass.label == 'brother'})
        assertTrue('duplicate target 2', dfc.targetElements.any {it.label == 'twin_sibling' && it.dataClass.label == 'brother'})
    }

    @Override
    protected ExcelDataFlowImporterParameters createImportParameters(Path path) throws IOException {
        ExcelDataFlowImporterParameters params = new ExcelDataFlowImporterParameters()
        FileParameter file = new FileParameter(path.toString(), 'application/vnd.ms-excel', Files.readAllBytes(path))

        params.setImportFile(file)
        params
    }

    @Override
    DataFlow saveDomain(DataFlow domain) {
        domain.save(validate: false, flush: true)
    }

    protected void importTestDataModels() {
        try {
            ExcelDataModelImporterService excelImporterService = getBean(ExcelDataModelImporterService)

            Path p = Paths.get('src/integration-test/resources/multiDataModelImport.xlsx')
            if (!Files.exists(p)) {
                fail('File multiDataModelImport.xlsx cannot be found')
            }
            ExcelFileImporterParameters params = new ExcelFileImporterParameters(finalised: false)
            FileParameter file = new FileParameter(p.toString(), 'application/vnd.ms-excel', Files.readAllBytes(p))

            params.setImportFile(file)
            List<DataModel> dataModels = excelImporterService.importDataModels(catalogueUser, params)
            dataModels*.save(flush: true)
        } finally {
            SessionFactory sessionFactory = getBean(SessionFactory)
            sessionFactory.getCurrentSession().clear()
        }
    }
}
