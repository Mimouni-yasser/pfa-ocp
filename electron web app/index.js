const { app, dialog, BrowserWindow } = require("electron")
let win;
function createWindow() {
 win = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
  });

  win.on('close', e => {
    e.preventDefault()
    dialog.showMessageBox({
      type: 'warning',
      buttons: ['save then exit', 'Exit'],
      cancelId: 1,
      defaultId: 0,
      title: 'Warning',
      detail: 'save changes before exiting?'
    }).then(({ response, checkboxChecked }) => {
      if (response) {
        win.destroy()
        app.quit()
      }
    })
  })
    win.loadFile('index.html');
}



app.whenReady().then( () => 
{
    createWindow();
})


