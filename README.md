# mdm-plugin-template

## How to use

* Clone this repository to a new folder
```bash
$ git clone git@github.com:MauroDataMapper-Plugins/mdm-plugin-template.git mdm-plugin-NAME_OF_PLUGIN
```
* Remove this section of the README and update `NAME_OF_PLUGIN` in the markdown template below
* Set `rootProject.name` in `settings.gradle` to `mdm-plugin-NAME_OF_PLUGIN`
* Create your code package, if the plugin will be published inside Mauro Data Mapper Plugins organisation then the package MUST be
`uk.ac.ox.softeng.maurodatamapper.plugins.your.plugin.name`
* Refactor `uk.ac.ox.softeng.maurodatamapper.plugins.template.TemplatePlugin` to be `YourPluginNamePlugin` in the correct package
* Update `src/main/resources/META-INF/services/uk.ac.ox.softeng.maurodatamapper.provider.plugin.MauroDataMapperPlugin` to hold the fully qualified
 name of your plugin class
* Write all your code and tests
* Declare your `mainClass` and `applicationScriptName` in `gradle.properties` (see other plugins for examples)

### AbstractMauroDataMapperPlugin

The following method implemented in your Plugin class allows you to wire any services/beans into the Grails application when it runs up.

```groovy
    @Override
    Closure doWithSpring() {
        {->
            
        }
    }
```

This should be used to add any services, such as Importers/Exporters, and make them available to the backend.
Do not implement any spring context configurations this can be done entirely using the method above.
The closure is passed to `grails.spring.BeanBuilder`.

### Markdown Template
```markdown
# mdm-plugin-NAME_OF_PLUGIN

| Branch | Build Status |
| ------ | ------------ |
| master | [![Build Status](https://jenkins.cs.ox.ac.uk/buildStatus/icon?job=Mauro+Data+Mapper+Plugins%2Fmdm-plugin-NAME_OF_PLUGIN%2Fmaster)](https://jenkins.cs.ox.ac.uk/blue/organizations/jenkins/Mauro%20Data%20Mapper%20Plugins%2Fmdm-plugin-NAME_OF_PLUGIN/branches) |
| develop | [![Build Status](https://jenkins.cs.ox.ac.uk/buildStatus/icon?job=Mauro+Data+Mapper+Plugins%2Fmdm-plugin-NAME_OF_PLUGIN%2Fdevelop)](https://jenkins.cs.ox.ac.uk/blue/organizations/jenkins/Mauro%20Data%20Mapper%20Plugins%2Fmdm-plugin-NAME_OF_PLUGIN/branches) |

## Requirements

* Java 12 (AdoptOpenJDK)
* Grails 4.0.3+
* Gradle 6.5+

All of the above can be installed and easily maintained by using [SDKMAN!](https://sdkman.io/install).
```
