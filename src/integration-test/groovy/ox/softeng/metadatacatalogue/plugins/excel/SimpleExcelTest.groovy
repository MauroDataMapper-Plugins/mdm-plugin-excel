package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.plugins.simpleexcel.SimpleExcelDataModelExporterService

import org.junit.Test
import static org.junit.Assert.assertEquals

class SimpleExcelTest {


    @Test
    void testSheetKey() {
        Map<String, String> results = [
            "My Model": "MM",
            "GEL: CancerSchema-v1.0.1": "GC101",
            "[GEL: CancerSchema-v1.0.1]": "GC101"
        ]

        results.each {entry ->
            assertEquals(SimpleExcelDataModelExporterService.createSheetKey(entry.key), entry.value)
        }



    }

}
