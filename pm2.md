# Start Node-RED automatically when Windows starts with PM2

To start Node-Red automatically, i found an interesting process magager tool for Node.js applications, pm2. 

The detailed infomation can be access here: https://github.com/Unitech/pm2 


### Installing on Windows7
```
$ npm i pm2 -g
```
### Necessary packages for auto starts
```
$ npm install pm2-windows-startup –g
$ npm i pm2-windows-service –g
$ pm2-service-install -n myservice
```
Then set the system variable：PM2_SERVICE_SCRIPTS = PM2

Restart myservice

### Node-red Start setting commands
```
$ pm2 start C:\Users\(Username)\AppData\Roaming\npm\node_modules\node-red\red.js
$ pm2 save
```
Note: _(Username)is your PC user id_. 

