const { app, BrowserWindow } = require('electron');
const path = require('path');

function createWindow () {
	const win = new BrowserWindow({
			width: 800,
			height: 600,
			webPreferences: {
				preload: path.join(__dirname, 'preload.js')
			}
	})

	win.loadFile('index.html')
}

app.whenReady().then(() => {
	createWindow()

	// MacOS - Open a window if none are open
	app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// Windows - Quit the app when all windows are closed
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})