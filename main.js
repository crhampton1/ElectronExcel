const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('path')
const util = require('util')
const fs = require('fs')
const excelToJson = require('convert-excel-to-json');


// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let win
const stat = util.promisify(fs.stat)
const DoTheExcel = util.promisify(excelToJson)

function createWindow () {
  // Create the browser window.
  win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  })

  const htmlPath = path.join('src','index.html')
  // and load the index.html of the app.
  win.loadFile(htmlPath)

  // Open the DevTools.
  win.webContents.openDevTools()

  // Emitted when the window is closed.
  win.on('closed', () => {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    win = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow)


// Quit when all windows are closed.
app.on('window-all-closed', () => {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (win === null) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.

ipcMain.on('files', async (event, filesArr) => {
  try{
    const data = await Promise.all(
      filesArr.map(async({pathName })=> ({
        ...await  DoTheExcelFunction(pathName)
      }))
    )
    win.webContents.send('metadata', data)
  } catch (error) {
    win.webContents.send('metadata:error', error)
  }
})

async function DoTheExcelFunction(pathName) {
  try {
   var JsonConverted = await excelToJson({
          source: fs.readFileSync(pathName),
          header:{
            rows: 1
        },
          columnToKey: {
            '*': '{{columnHeader}}'
          } 
        })
    
   
    var data = []
    
     
    for await (eachField of JsonConverted.ReportData){
      var newObject = {
        Church: '',
        Organization: '',
        People: ''
      }
      

      if(eachField.Id.includes('C')){
        newObject.Church = eachField.Id.substring(0, eachField.Id.length -1) 
      } else if (eachField.Id.includes('O')){
        newObject.Organization = eachField.Id.substring(0, eachField.Id.length -1) 
      } else if (eachField.Id.includes('I')){
        newObject.People = eachField.Id.substring(0, eachField.Id.length -1) 
      } else {
        console.log("NULL")
        
      }

      data.push(newObject)
    }
   
    var bigData = {data}

    
    return bigData

  } catch (error) {
    win.webContents.send('metadata:error', error)
  }
} 