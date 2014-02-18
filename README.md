excel-integration-example
=========================

Example of Excel integration with OpenFin Desktop

This package contains 2 components, Excel Add-in and a web app.  The web app, in web directory, runs in OpenFin Desktop environment and pubishes current timestamp once per second.  The Excel Add-in, contained in FinDesktopAddin-packed.xll, enables Excel to subscribe to messages from the web app and display the same timestamp published by the web app.

To run the example:

1. Install content of web directory to a web server so publish.html can be accessed via an URL.
2. Start OpenFin Desktop
3. Create an app with URL from step 1 in OpenFin Desktop, and start it.  The page should shows that curremt timestamp is being published once per second.  It also shows an App Id, which is needed in step 5.
4. Start Excel
5. Load FinDesktopAddin-packed.xll in Excel.
6. In any cell of Excel, enter function: =FinDesktop("publisher_app_id", "ExcelData", "timestamp").  publish_app_id is from step 3.  The cell should show same timestamp as publish app.


If you have questions, please contact us at support@openfin.co
