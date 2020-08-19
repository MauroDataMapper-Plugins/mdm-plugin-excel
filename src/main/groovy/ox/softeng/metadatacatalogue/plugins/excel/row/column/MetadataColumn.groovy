package ox.softeng.metadatacatalogue.plugins.excel.row.column

/**
 * @since 13/03/2018
 */
class MetadataColumn {

    String namespace
    String key
    String value

    boolean equals(o) {
        if (this.is(o)) return true
        if (getClass() != o.class) return false

        MetadataColumn that = (MetadataColumn) o

        if (key != that.key) return false
        if (namespace != that.namespace) return false

        return true
    }

    int hashCode() {
        int result
        result = (namespace != null ? namespace.hashCode() : 0)
        result = 31 * result + (key != null ? key.hashCode() : 0)
        return result
    }

    String toString() {
        "${namespace}/${key}"
    }
}
