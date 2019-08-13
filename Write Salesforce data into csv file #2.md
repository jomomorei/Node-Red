## How to write into CSV file 

In the part #1 the construction of excel output flow has been introduced. 
The big difference between them is the output data format: csv node is target for string while excel node can process json data.

The image is as same as in the part #1.
![Complete Image](/img/b6NDwE.png)

Now the foucs is the below banch.

#### 1.set msg.payload

![Complete Image](/img/7yMB63.png)

This is a change node. The data we get on the last node(salesforce connection, a soql node) is a json object. It contains two attributions: size and records. By this node, we persist records only.

#### 2.json to csv converter

This is a key step but simple step. It will convert the json data to string form. No parameter is needed. 

#### 3. write into csv

![Complete Image](/img/2dfAMQ.png)

This is a file node. File nodes have two pattern, in and out. Here is the out one. Open the sidebar, specify the file path and name. Three anction modes can be selected: append, overwrite and delete.

In fact, i didn't use csv node in this branch flow. Instead i used the json-to-csv-converter node because of its simplicity.
But if you want to specify the separator character, try csv node.

