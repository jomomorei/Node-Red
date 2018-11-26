
# Downloading Salesforce Data Writing Into Excel/CSV File 

I happend accross Node-Red months ago. It's a programming tool for wiring together hardware devices, APIs and online services in new and interesting ways. 
The URL is here: https://nodered.org/

Last week, i tried to use it to connnect to salesforce and get the data by SOQL. Then saved the data into Excel file and CSV files. A little pity, japanese can not be recognized in the CSV file.

![Complete Image](/img/b6NDwE.png)

## Node-Red For Salesforce

* Add an Inject Node

Drag inject node onto flow workspace from palette. It allow you to inject messages into a flow, either by clicking the button on the node, or setting a time interval between injects.

If you want to change the inject message, open the sidebar. 

![Complete Image](/img/KXlvz4.png)

Here I set the Payload（ペイロード）is time form（日時）, Topic（トピック）is blank. The Repeat（繰り返し）can be set dependent on the frequency you want it to triggle. The Name（名前）can be blank. The default value is timestamp.

* Add an Soql Node

This Node is for executing a SOQL query. SOQL is short for Salesforce Object Query Language, to search your organization’s Salesforce data for specific information. More detailed information can be get here: https://developer.salesforce.com/docs/atlas.en-us.soql_sosl.meta/soql_sosl/sforce_api_calls_soql.htm .

The first time to use it, you should download the node firstly. Click the top-right icon, select the Palette Manager（パレットの管理）. In the Add Node tab（ノードを追加）, search "salesforce" and add it as below image. Then you can find several nodes in the Salesforce sidebar tab. Drag Soql node onto workspace. 

![Complete Image](/img/njL2qR.png)
![Complete Image](/img/j6swxV.png)

