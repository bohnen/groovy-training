/**
 * Download smartsheet project file as xls and format to print.
 */

@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
                @Grab(group='commons-io', module='commons-io', version='2.4'),
                @Grab(group='org.codehaus.groovy.modules.http-builder', module='http-builder', version='0.7.1' )

        ]
)
import org.apache.poi.ss.usermodel.*
import groovyx.net.http.HTTPBuilder
import groovyx.net.http.Method

// input properties
def prop = new ConfigSlurper().parse(new File(".smartsheet.groovy").toURI().toURL())

// output directory and file
new File("out").mkdir();
def file = new File("./out/out.xls")

def http = new HTTPBuilder('https://api.smartsheet.com')
http.request(Method.GET){ req ->
    uri.path = prop.path
    headers.Authorization = "Bearer ${prop.apikey}"
    headers.Accept = 'application/vnd.ms-excel'

    response.success = {resp, reader ->
        file << reader
    }
}
