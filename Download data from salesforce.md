
# Downloading Salesforce Data Writing Into EXCEL File 

I happend accross Node-Red months ago. It's a programming tool for wiring together hardware devices, APIs and online services in new and interesting ways. 
The URL is here: https://nodered.org/

Last week, i tried to use it to connnect to salesforce and get the data by SOQL. Then saved the data into Excel file and CSV files. A little pity, japanese can not be recognized in the CSV file.

![Complete Image](/img/b6NDwE.png)

## Node-Red For Salesforce

* Add an Inject Node

Drag inject node onto flow workspace from palette. It allow you to inject messages into a flow, either by clicking the button on the node, or setting a time interval between injects.

If you want to change the inject message, open the sidebar. 

![Complete Image](/img/KXlvz4.png)

Here I set the Payload（ペイロード）is time form（日時）, Topic（トピック）is blank. The recurring（繰り返し）can be set dependent on the frequency you want it to triggle. The Name（名前）can be blank. The default value is timestamp.

* Add an Soql Node

This Node is for executing a SOQL query. SOQL is short for Salesforce Object Query Language, to search your organization’s Salesforce data for specific information. More detailed information can be get here: https://developer.salesforce.com/docs/atlas.en-us.soql_sosl.meta/soql_sosl/sforce_api_calls_soql.htm .

The first time to use it, you should download the node firstly. Click the top-right icon, select the Palette Manager（パレットの管理）. In the Add-Node tab（ノードを追加）, search "salesforce" and add it as below image. Then you can find several nodes in the Salesforce sidebar tab. Drag Soql node onto workspace. 

![Complete Image](/img/njL2qR.png)

![Complete Image](/img/j6swxV.png)

Double click the Soql node, the sidebar will come out. Here, editing the name and valid query. 
![Complete Image](/img/LuZTpn.png)

Click the pen icon, in this page, the key information connecting to salesforce is required.

![Complete Image](/img/EdsgKk.png)

The Name can be arbitrary. Environment is Production. Username is salesforce admin ID. And the Password is combinations of security token and salesforce password. Then turn into salesforce page and we will obtain the left three items: Consumer Key, Consumer Secret and Callback URL.

Get into Setting page. In the left search box, inputting "app", find "Application" and click it. On the right-bottom page, new an application(the first time). 

Input some key information to set the new application connection. As below:(ja)
![Complete Image](/img/Xq4kaW.png)

Finish version is like this. Consumer Key and Consumer Secret have be produced automatically.   
![Complete Image](/img/RiNCQT.png)

Copy the three items and back to the node-red page, fill them and soql node setting is finished.

* Add a function node

Drag function node onto flow workspace. And open the sidebar.

In the soql node, we fetch data from salesforce in json form. Although, the next excel node can converts a JSON or array of JSONs to excel and writes to the file. The following coding is necessary to fetch clean data.

```javascript
var str=msg.payload.records;
var astr=new Array();
for(i=0;i<str.length;i++){
    var p=eval('('+JSON.stringify(str[i])+')');
    astr[i]=p;
}
msg.payload=astr;
return msg;
```

* Add a EXCEL node

Drag excel node to the behind, and specify the absolute filepath.

* Add debug node

If you want to watch the output msg in the console, add a debug node. 

Finally, Wire nodes like the top image and click the Payload button on the top-right. The Excel export setup has been done.

If you want to make it work immediately, click the inject button. And you will see the output data in the debug tab in the console.



