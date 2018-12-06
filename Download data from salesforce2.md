## How to write into CSV file 

In 1 the construction of excel output flow is introduced. 
The big difference between them is the data format that csv node is target for string while excel node can process json data.

The image is as same as in 1.
![Complete Image](/img/b6NDwE.png)

The foucs is the below banch.

#### 1.set msg.payload

![Complete Image](/img/7yMB63.png)

This is a change node. The data we get on the last node(salesforce connection, a soql node) is a json object. It contains two attributions: size and records. By this node, we persist records only.

#### 2.json to csv converter

This is a key step as well as a simple step. As you shoud do nothing, just drag this node onto the workspace. No parameter is needed.
It will convert json data to string form.

#### 3. write into csv

![Complete Image](/img/2dfAMQ.png)

This is a file node. File nodes have two pattern: in and out. Of course, here is the out one. Open the sidebar, file path and name should be specified. It has three anction modes: append, overwrite and delete.

In fact, i didn't use csv node in this branch flow but used the json-to-csv-converter node. Because it's more simple.
But if you want to specify the separator character, try csv node.

