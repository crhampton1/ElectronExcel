const {ipcRenderer} = require('electron')
var electron = require('electron').remote;
var XLSX = require('xlsx');



const submitListner = document
    .querySelector('form')
    .addEventListener('submit', (event) => {
        event.preventDefault()

        const files = [...document.getElementById('filePicker').files]
        const filesFormated = files.map(({name, path: pathName}) => ({
            name,
            pathName
        }))
        console.log(filesFormated)
        ipcRenderer.send('files', filesFormated)
    })

    ipcRenderer.on('metadata', (event, response) => {
        const pre = document.getElementById('data')

        pre.innerText = response 
    })

    ipcRenderer.on('metadata:error', (event, error) => {
        console.error(error)
    });

    